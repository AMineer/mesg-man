Describe 'mesg-man script tests' {
    
    It 'should handle errors gracefully' {
        { .\mesg-man.ps1 -InvalidParameter } | Should -Throw
    }

    It 'should run in WhatIf mode without errors' {
        { .\mesg-man.ps1 -WhatIf } | Should -Not -Throw
    }
}