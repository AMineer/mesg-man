# mesg-man
Mail Enabled Security Group console tool

---

## Features

- Add users to groups based on input from a CSV file.
- Supports `-WhatIf` mode to simulate changes without making actual modifications.
- Logs actions and errors for better traceability.
- Cleans up connections after execution.

---

## Prerequisites

- PowerShell 5.1 or later.
- Exchange Online Management Module (if applicable).
- Pester (for running tests).

To install Pester, run:
```powershell
Install-Module -Name Pester -Force -SkipPublisherCheck
```

---

## Usage

1. **Run the script**:
   ```powershell
   .\mesg-man.ps1 -Parameter1 Value1 -Parameter2 Value2
   ```

2. **Run in `-WhatIf` mode**:
   ```powershell
   .\mesg-man.ps1 -WhatIf
   ```

3. **Handle errors gracefully**:
   The script is designed to log errors and exit gracefully if invalid parameters are provided.

---

## Testing with Pester

The script includes a test suite written in Pester. The tests are located in the 

Tests

 folder.

### Running Tests

To run the tests, execute the following command:
```powershell
Invoke-Pester -Path .\Tests\mesg-man.Tests.ps1
```

### Test Cases

- **Error Handling**: Verifies that the script throws an error when invalid parameters are provided.
- **WhatIf Mode**: Ensures the script runs in `-WhatIf` mode without throwing any errors.

Example test file (`mesg-man.Tests.ps1`):
```powershell
Describe 'mesg-man script tests' {
    It 'should handle errors gracefully' {
        { .\mesg-man.ps1 -InvalidParameter } | Should -Throw
    }

    It 'should run in WhatIf mode without errors' {
        { .\mesg-man.ps1 -WhatIf } | Should -Not -Throw
    }
}
```

---

## Logging

The script uses a `Write-Log` function to log actions and errors. Logs are categorized as `INFO`, `ERROR`, etc., for better readability.

---

## Cleanup

The script ensures that connections (e.g., Exchange Online) are cleaned up after execution:
```powershell
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Log "Disconnected from Exchange Online." 'INFO'
```

---

## Contributing

Feel free to submit issues or pull requests to improve the script or add new features.

---

## License

This project is licensed under the MIT License.