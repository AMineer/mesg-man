[CmdletBinding()]
param(
    [switch]$WhatIf
)

<#
.SYNOPSIS
    Add users to a mail-enabled security group (or any distribution group) in Exchange Online.

.DESCRIPTION
    - Ensures ExchangeOnlineManagement module is installed
    - Connects to Exchange Online
    - Presents a menu:
        1) Single user
        2) Multiple users (comma-separated)
        3) CSV file
        4) View current group members
    - Skips users who are already members (logs + writes to console)
    - Logs all activity to a timestamped log file in your Documents folder
    - Supports -WhatIf switch to simulate without doing any changes

.NOTES
    Run in pwsh / PowerShell:
        pwsh ./MESG-Man.ps1
        pwsh ./MESG-Man.ps1 -WhatIf
#>

#-----------------------------------------------------------
# Global logging setup
#-----------------------------------------------------------

$documentsFolder = [Environment]::GetFolderPath('MyDocuments')
if (-not $documentsFolder) {
    $documentsFolder = (Get-Location).Path
}

$script:LogPath = Join-Path -Path $documentsFolder -ChildPath ("MESG_AddMembers_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO','WARN','ERROR','WHATIF')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$timestamp [$Level] $Message"

    # Write to log file
    try {
        Add-Content -Path $script:LogPath -Value $line -ErrorAction SilentlyContinue
    }
    catch {
        # If logging fails, just ignore (still write to console)
    }

    # Pick a console color
    switch ($Level) {
        'INFO'   { $color = 'White' }
        'WARN'   { $color = 'Yellow' }
        'ERROR'  { $color = 'Red' }
        'WHATIF' { $color = 'Cyan' }
    }

    Write-Host $line -ForegroundColor $color
}

#-----------------------------------------------------------
# Fancy banner / menu header (MESG-MAN)
#-----------------------------------------------------------

function Show-Banner {
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Magenta
    Write-Host "                      MESG-MAN" -ForegroundColor Magenta
    Write-Host "          Mail-Enabled Security Group Manager" -ForegroundColor Magenta
    Write-Host "======================================================" -ForegroundColor Magenta
    Write-Host "" -ForegroundColor Magenta
    Write-Host "                 M E S G  -  M A N" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "Log file: $script:LogPath" -ForegroundColor DarkGray

    if ($WhatIf) {
        Write-Host ">>> WHAT-IF MODE ENABLED: No changes will be made. <<<" -ForegroundColor Cyan
    }

    Write-Host ""
}

#-----------------------------------------------------------
# Ensure ExchangeOnlineManagement module
#-----------------------------------------------------------

function Install-ExchangeOnlineModule {
    $moduleName = 'ExchangeOnlineManagement'

    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Log "[$moduleName] module not found. Installing for current user..." 'WARN'
        try {
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "[$moduleName] installed successfully." 'INFO'
        }
        catch {
            Write-Log "Failed to install $moduleName : $_" 'ERROR'
            throw
        }
    }
    else {
        Write-Log "[$moduleName] module is already available." 'INFO'
    }

    Import-Module $moduleName -ErrorAction Stop
    Write-Log "[$moduleName] module imported." 'INFO'
}

#-----------------------------------------------------------
# Connect to Exchange Online
#-----------------------------------------------------------

function Connect-ToExchangeOnline {
    try {
        Write-Log "Connecting to Exchange Online..." 'INFO'
        Connect-ExchangeOnline -ErrorAction Stop
        Write-Log "Connected to Exchange Online." 'INFO'
    }
    catch {
        Write-Log "Failed to connect to Exchange Online: $_" 'ERROR'
        throw
    }
}

#-----------------------------------------------------------
# Get group members into a hash set
#-----------------------------------------------------------

function Get-GroupAndMemberSet {
    param(
        [Parameter(Mandatory)]
        [string]$GroupIdentity
    )

    try {
        Write-Log "Retrieving members of group: $GroupIdentity" 'INFO'
        $existingMembers = Get-DistributionGroupMember -Identity $GroupIdentity -ResultSize Unlimited -ErrorAction Stop
    }
    catch {
        Write-Log "Failed to get members for group [$GroupIdentity]: $_" 'ERROR'
        throw
    }

    $memberSet = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($m in $existingMembers) {
        $addresses = @()

        if ($m.PrimarySmtpAddress)  { $addresses += $m.PrimarySmtpAddress.ToString() }
        if ($m.WindowsEmailAddress) { $addresses += $m.WindowsEmailAddress.ToString() }

        foreach ($addr in $addresses) {
            [void]$memberSet.Add($addr.ToLower())
        }
    }

    Write-Log "Loaded $($memberSet.Count) existing members for group [$GroupIdentity]." 'INFO'

    return [PSCustomObject]@{
        GroupIdentity = $GroupIdentity
        MemberSet     = $memberSet
    }
}

#-----------------------------------------------------------
# Add a single user to group (with membership check & WhatIf)
#-----------------------------------------------------------

function Add-UserToGroup {
    param(
        [Parameter(Mandatory)]
        [string]$UserIdentifier,

        [Parameter(Mandatory)]
        [string]$GroupIdentity,

        [Parameter(Mandatory)]
        [System.Collections.Generic.HashSet[string]]$MemberSet,

        [switch]$WhatIf
    )

    $user = $UserIdentifier.Trim()
    if (-not $user) { return }

    $key = $user.ToLower()

    if ($MemberSet.Contains($key)) {
        Write-Log "[$user] is already a member of [$GroupIdentity]. Skipping." 'WARN'
        return
    }

    if ($WhatIf) {
        Write-Log "WHAT-IF: Would add [$user] to [$GroupIdentity]." 'WHATIF'
        return
    }

    try {
        Add-DistributionGroupMember -Identity $GroupIdentity -Member $user -ErrorAction Stop
        Write-Log "Added [$user] to [$GroupIdentity]." 'INFO'
        [void]$MemberSet.Add($key)  # track new member for this run
    }
    catch {
        Write-Log "Failed to add [$user] to [$GroupIdentity]: $_" 'ERROR'
    }
}

#-----------------------------------------------------------
# Get users from CSV
#-----------------------------------------------------------

function Get-UsersFromCsv {
    param(
        [Parameter(Mandatory)]
        [string]$CsvPath
    )

    if (-not (Test-Path -Path $CsvPath)) {
        Write-Log "CSV file not found at: $CsvPath" 'ERROR'
        return @()
    }

    try {
        $rows = Import-Csv -Path $CsvPath -ErrorAction Stop
    }
    catch {
        Write-Log "Failed to read CSV [$CsvPath]: $_" 'ERROR'
        return @()
    }

    if (-not $rows -or $rows.Count -eq 0) {
        Write-Log "CSV file [$CsvPath] is empty." 'WARN'
        return @()
    }

    $candidateColumns = @(
        'UserPrincipalName',
        'UPN',
        'Email',
        'Mail',
        'PrimarySmtpAddress'
    )

    $firstRow = $rows[0]
    $availableColumns = $firstRow.PSObject.Properties.Name

    $chosenColumn = $candidateColumns | Where-Object { $_ -in $availableColumns } | Select-Object -First 1

    if (-not $chosenColumn) {
        Write-Log "Could not find a suitable email column in CSV [$CsvPath]." 'ERROR'
        Write-Log "Expected one of: $($candidateColumns -join ', ')" 'WARN'
        Write-Log "Found columns: $($availableColumns -join ', ')" 'WARN'
        return @()
    }

    Write-Log "Using column [$chosenColumn] from CSV [$CsvPath] for user identifiers." 'INFO'

    $users = @()
    foreach ($row in $rows) {
        $value = $row.$chosenColumn
        if ($value) { $users += $value }
    }

    Write-Log "Extracted $($users.Count) user(s) from CSV." 'INFO'
    return $users
}

#-----------------------------------------------------------
# MAIN
#-----------------------------------------------------------

Write-Log "=== MESG-MAN Tool started ===" 'INFO'
Write-Log "Log file path: $script:LogPath" 'INFO'
if ($WhatIf) {
    Write-Log "WHAT-IF MODE: No changes will be made to Exchange Online." 'WHATIF'
}

Show-Banner

try {
    # 1. Ensure module
    Install-ExchangeOnlineModule

    # 2. Connect to Exchange Online
    Connect-ToExchangeOnline

    # 3. Prompt for group
    $groupIdentity = Read-Host "Enter the mail-enabled security group (name, alias, or email address)"

    if (-not $groupIdentity) {
        Write-Log "No group specified. Exiting." 'ERROR'
        return
    }

    Write-Log "Target group: [$groupIdentity]" 'INFO'

    # 4. Get existing members + hash set
    $groupInfo = Get-GroupAndMemberSet -GroupIdentity $groupIdentity
    $memberSet = $groupInfo.MemberSet

    # 5. Show menu
    Write-Host ""
    Write-Host "Select input method:" -ForegroundColor Cyan
    Write-Host "  1) Single user"
    Write-Host "  2) Multiple users (comma-separated)"
    Write-Host "  3) CSV file"
    Write-Host "  4) View current group members"
    Write-Host ""

    $choice = Read-Host "Enter 1, 2, 3, 4, or 5"
    Write-Log "Menu choice: $choice" 'INFO'

    switch ($choice) {
        '1' {
            # Single user
            $user = Read-Host "Enter the user's email address / UPN"
            if ($user) {
                Add-UserToGroup -UserIdentifier $user -GroupIdentity $groupIdentity -MemberSet $memberSet -WhatIf:$WhatIf
            }
            else {
                Write-Log "No user provided for single-user mode. Nothing to do." 'WARN'
            }
        }

        '2' {
            # Multiple users comma-separated
            $raw = Read-Host "Enter email addresses / UPNs separated by commas"
            if ($raw) {
                $users = $raw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                Write-Log "Parsed $($users.Count) user(s) from comma-separated input." 'INFO'
                foreach ($u in $users) {
                    Add-UserToGroup -UserIdentifier $u -GroupIdentity $groupIdentity -MemberSet $memberSet -WhatIf:$WhatIf
                }
            }
            else {
                Write-Log "No users provided for multi-user mode. Nothing to do." 'WARN'
            }
        }

        '3' {
            # CSV
            $csvPath = Read-Host "Enter full path to CSV file"
            $usersFromCsv = Get-UsersFromCsv -CsvPath $csvPath

            if ($usersFromCsv.Count -eq 0) {
                Write-Log "No users found in CSV mode. Nothing to do." 'WARN'
            }
            else {
                foreach ($u in $usersFromCsv) {
                    Add-UserToGroup -UserIdentifier $u -GroupIdentity $groupIdentity -MemberSet $memberSet -WhatIf:$WhatIf
                }
            }
        }

        '4' {
            # View current group members
            Write-Host "Current group members:" -ForegroundColor Green
            $groupInfo.MemberSet | ForEach-Object { Write-Host $_ }
            Write-Host ""
            Write-Host "Press Enter to return to the exit..." -ForegroundColor Cyan
            Read-Host
        }

        Default {
            Write-Log "Invalid choice [$choice]. Exiting." 'ERROR'
        }
    }

}
finally {
    # 6. Clean up connection
    try {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        Write-Log "Disconnected from Exchange Online." 'INFO'
    }
    catch {
        Write-Log "Error during Disconnect-ExchangeOnline (ignored): $_" 'WARN'
    }
}
