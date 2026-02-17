# Import module before tests
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath '..\ReverseDSC.Core.psm1'
Import-Module -Name $modulePath -Force

InModuleScope 'ReverseDSC.Core' {
    Describe 'ConvertTo-EscapedDSCString' {
    Context 'When the input string is null or empty' {
        It 'Should return the same empty string' {
            $result = ConvertTo-EscapedDSCString -InputString ''
            $result | Should -BeNullOrEmpty
        }

        It 'Should return null when passed null' {
            $result = ConvertTo-EscapedDSCString -InputString $null
            $result | Should -BeNullOrEmpty
        }
    }

    Context 'When the input string contains backticks' {
        It 'Should escape backticks by doubling them' {
            $result = ConvertTo-EscapedDSCString -InputString 'Hello`World'
            $result | Should -Be 'Hello``World'
        }
    }

    Context 'When the input string contains dollar signs' {
        It 'Should escape dollar signs by default' {
            $result = ConvertTo-EscapedDSCString -InputString 'Price is $100'
            $result | Should -Be 'Price is `$100'
        }

        It 'Should preserve dollar signs when AllowVariables is specified' {
            $result = ConvertTo-EscapedDSCString -InputString 'Value is $var' -AllowVariables
            $result | Should -Be 'Value is $var'
        }
    }

    Context 'When the input string contains European quotation marks' {
        It 'Should escape U+201E (double low-9 quotation mark)' {
            $input201E = "test$([char]0x201E)value"
            $result = ConvertTo-EscapedDSCString -InputString $input201E
            $result | Should -Be "test``$([char]0x201E)value"
        }

        It 'Should escape U+201C (left double quotation mark)' {
            $input201C = "test$([char]0x201C)value"
            $result = ConvertTo-EscapedDSCString -InputString $input201C
            $result | Should -Be "test``$([char]0x201C)value"
        }

        It 'Should escape U+201D (right double quotation mark)' {
            $input201D = "test$([char]0x201D)value"
            $result = ConvertTo-EscapedDSCString -InputString $input201D
            $result | Should -Be "test``$([char]0x201D)value"
        }
    }

    Context 'When the input string contains double quotes' {
        It 'Should escape double quotes' {
            $result = ConvertTo-EscapedDSCString -InputString 'She said "hello"'
            $result | Should -Be 'She said `"hello`"'
        }
    }

    Context 'When the input string contains double quotes and escape characters' {
        It 'Should escape double quotes and escape characters' {
            $result = ConvertTo-EscapedDSCString -InputString 'She said "hello" with `"Escaped Text`"'
            $result | Should -Be 'She said `"hello`" with ```"Escaped Text```"'
        }
    }

    Context 'When the input string is plain text without special characters' {
        It 'Should return the string unchanged' {
            $result = ConvertTo-EscapedDSCString -InputString 'Normal text'
            $result | Should -Be 'Normal text'
        }
    }
}

Describe 'ConvertTo-DSCStringValue' {
    Context 'When the value is null' {
        It 'Should return empty double-quoted string' {
            $result = ConvertTo-DSCStringValue -Value $null
            $result | Should -Be '""'
        }
    }

    Context 'When NoEscape is true' {
        It 'Should return the raw value without escaping' {
            $result = ConvertTo-DSCStringValue -Value 'MyValue' -NoEscape $true
            $result | Should -Be 'MyValue'
        }
    }

    Context 'When NoEscape is false (default)' {
        It 'Should return the value wrapped in double quotes' {
            $result = ConvertTo-DSCStringValue -Value 'SimpleString'
            $result | Should -Be '"SimpleString"'
        }

        It 'Should escape special characters in the value' {
            $result = ConvertTo-DSCStringValue -Value 'Value with $var'
            $result | Should -Be '"Value with `$var"'
        }
    }

    Context 'When AllowVariables is true' {
        It 'Should preserve dollar signs in the value' {
            $result = ConvertTo-DSCStringValue -Value 'Value with $var' -AllowVariables $true
            $result | Should -Be '"Value with $var"'
        }
    }
}

Describe 'ConvertTo-DSCBooleanValue' {
    It 'Should return $True for true values' {
        $result = ConvertTo-DSCBooleanValue -Value $true
        $result | Should -Be '$True'
    }

    It 'Should return $False for false values' {
        $result = ConvertTo-DSCBooleanValue -Value $false
        $result | Should -Be '$False'
    }
}

Describe 'ConvertTo-DSCCredentialValue' {
    Context 'When the value is null' {
        It 'Should return a Get-Credential command with the parameter name' {
            $result = ConvertTo-DSCCredentialValue -Value $null -ParameterName 'Credential'
            $result | Should -Be 'Get-Credential -Message Credential'
        }
    }

    Context 'When the value is a PSCredential with a UPN username' {
        BeforeAll {
            $securePassword = ConvertTo-SecureString -String 'Password123' -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential ('admin@contoso.com', $securePassword)
        }

        It 'Should return a $Creds variable based on the username part' {
            $result = ConvertTo-DSCCredentialValue -Value $credential -ParameterName 'Credential'
            $result | Should -Be '$Credsadmin'
        }
    }

    Context 'When the value is a PSCredential with a domain\user username' {
        BeforeAll {
            $securePassword = ConvertTo-SecureString -String 'Password123' -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential ('CONTOSO\admin', $securePassword)
        }

        It 'Should return a $Creds variable based on the username after backslash' {
            $result = ConvertTo-DSCCredentialValue -Value $credential -ParameterName 'Credential'
            $result | Should -Be '$Credsadmin'
        }
    }

    Context 'When the value is a PSCredential with special characters in username' {
        BeforeAll {
            $securePassword = ConvertTo-SecureString -String 'Password123' -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential ('CONTOSO\admin-user.name', $securePassword)
        }

        It 'Should sanitize special characters in the variable name' {
            $result = ConvertTo-DSCCredentialValue -Value $credential -ParameterName 'Credential'
            $result | Should -Be '$Credsadmin_user_name'
        }
    }
}

Describe 'ConvertTo-DSCHashtableValue' {
    It 'Should format a single-entry hashtable correctly' {
        $hashtable = @{ Key1 = 'Value1' }
        $result = ConvertTo-DSCHashtableValue -Value $hashtable
        $result | Should -BeLike '@{*Key1 = "Value1"*}'
    }

    It 'Should format a multi-entry hashtable correctly' {
        $hashtable = [ordered]@{ Key1 = 'Value1'; Key2 = 'Value2' }
        $result = ConvertTo-DSCHashtableValue -Value $hashtable
        $result | Should -BeLike '@{Key1*Key2*}'
        $result | Should -Match 'Key1 = "Value1"'
        $result | Should -Match 'Key2 = "Value2"'
    }

    It 'Should wrap the result in @{ }' {
        $hashtable = @{ A = 'B' }
        $result = ConvertTo-DSCHashtableValue -Value $hashtable
        $result | Should -Match '^@\{'
        $result | Should -Match '\}$'
    }
}

Describe 'ConvertTo-DSCStringArrayValue' {
    Context 'When the value is null or empty' {
        It 'Should return @() for null value' {
            $result = ConvertTo-DSCStringArrayValue -Value $null
            $result | Should -Be '@()'
        }

        It 'Should return @() for empty array' {
            $result = ConvertTo-DSCStringArrayValue -Value @()
            $result | Should -Be '@()'
        }
    }

    Context 'When the value is a single-element array' {
        It 'Should return a properly formatted array string' {
            $result = ConvertTo-DSCStringArrayValue -Value @('Item1')
            $result | Should -Be '@("Item1")'
        }
    }

    Context 'When the value is a multi-element array' {
        It 'Should return a comma-separated array string' {
            $result = ConvertTo-DSCStringArrayValue -Value @('Item1', 'Item2', 'Item3')
            $result | Should -Be '@("Item1","Item2","Item3")'
        }
    }

    Context 'When NoEscape is true' {
        It 'Should not escape special characters in array elements' {
            $result = ConvertTo-DSCStringArrayValue -Value @('$var1', '$var2') -NoEscape $true
            $result | Should -Be '@("$var1","$var2")'
        }
    }
}

Describe 'ConvertTo-DSCIntegerArrayValue' {
    Context 'When the value is null or empty' {
        It 'Should return @() for null value' {
            $result = ConvertTo-DSCIntegerArrayValue -Value $null
            $result | Should -Be '@()'
        }

        It 'Should return @() for empty array' {
            $result = ConvertTo-DSCIntegerArrayValue -Value @()
            $result | Should -Be '@()'
        }
    }

    Context 'When the value contains integers' {
        It 'Should return a comma-separated integer array' {
            $result = ConvertTo-DSCIntegerArrayValue -Value @(1, 2, 3)
            $result | Should -Be '@(1,2,3)'
        }

        It 'Should handle a single integer' {
            $result = ConvertTo-DSCIntegerArrayValue -Value @(42)
            $result | Should -Be '@(42)'
        }
    }
}

Describe 'ConvertTo-DSCObjectArrayValue' {
    Context 'When the value is null or empty' {
        It 'Should return @() for null value' {
            $result = ConvertTo-DSCObjectArrayValue -Value $null
            $result | Should -Be '@()'
        }

        It 'Should return @() for empty array' {
            $result = ConvertTo-DSCObjectArrayValue -Value @()
            $result | Should -Be '@()'
        }
    }

    Context 'When the value contains strings' {
        It 'Should format string elements with quotes' {
            $result = ConvertTo-DSCObjectArrayValue -Value @('A', 'B', 'C')
            $result | Should -Be '@("A","B","C")'
        }
    }

    Context 'When the value contains hashtables' {
        It 'Should format each hashtable in the array' {
            $value = @(
                @{ Name = 'Item1' }
            )
            $result = ConvertTo-DSCObjectArrayValue -Value $value
            $result | Should -BeLike '@(@{*Name*Item1*})'
        }

        It 'Should handle null values in hashtable entries' {
            $value = @(
                @{ Name = $null }
            )
            $result = ConvertTo-DSCObjectArrayValue -Value $value
            $result | Should -Match '\$null'
        }

        It 'Should handle array values in hashtable entries' {
            $value = @(
                @{ Items = @('A', 'B') }
            )
            $result = ConvertTo-DSCObjectArrayValue -Value $value
            $result | Should -Match "@\("
        }
    }

    Context 'When NoEscape is true' {
        It 'Should not escape string values' {
            $result = ConvertTo-DSCObjectArrayValue -Value @('$var') -NoEscape $true
            $result | Should -Match '\$var'
        }
    }
}

Describe 'Get-DSCDependsOnBlock' {
    It 'Should generate a proper DependsOn clause for a single dependency' {
        $result = Get-DSCDependsOnBlock -DependsOnItems @('[xWebsite]DefaultSite')
        $result | Should -Be '@("[xWebsite]DefaultSite");'
    }

    It 'Should generate a proper DependsOn clause for multiple dependencies' {
        $result = Get-DSCDependsOnBlock -DependsOnItems @('[xWebsite]DefaultSite', '[xSPSite]MainSite')
        $result | Should -Be '@("[xWebsite]DefaultSite","[xSPSite]MainSite");'
    }
}

Describe 'Save-Credentials' {
    BeforeEach {
        # Reset the credentials repo before each test
        $Script:CredsRepo = @()
    }

    It 'Should add a new username to the credentials repository' {
        Save-Credentials -UserName 'CONTOSO\admin'
        $Script:CredsRepo | Should -Contain 'contoso\admin'
    }

    It 'Should store usernames in lowercase' {
        Save-Credentials -UserName 'CONTOSO\ADMIN'
        $Script:CredsRepo | Should -Contain 'contoso\admin'
    }

    It 'Should not duplicate usernames' {
        Save-Credentials -UserName 'CONTOSO\admin'
        Save-Credentials -UserName 'contoso\admin'
        $Script:CredsRepo | Should -HaveCount 1
    }
}

Describe 'Get-Credentials' {
    BeforeAll {
        $Script:CredsRepo = @()
        Save-Credentials -UserName 'CONTOSO\admin'
    }

    It 'Should return the username when it exists in the repository' {
        $result = Get-Credentials -UserName 'CONTOSO\admin'
        $result | Should -Be 'contoso\admin'
    }

    It 'Should return null when the username is not in the repository' {
        $result = Get-Credentials -UserName 'CONTOSO\nonexistent'
        $result | Should -BeNullOrEmpty
    }
}

Describe 'Test-Credentials' {
    BeforeAll {
        $Script:CredsRepo = @()
        Save-Credentials -UserName 'CONTOSO\admin'
    }

    It 'Should return true when the username exists' {
        $result = Test-Credentials -UserName 'CONTOSO\admin'
        $result | Should -BeTrue
    }

    It 'Should return false when the username does not exist' {
        $result = Test-Credentials -UserName 'CONTOSO\unknown'
        $result | Should -BeFalse
    }
}

Describe 'Resolve-Credentials' {
    It 'Should return $Creds<username> for domain\user format' {
        $result = Resolve-Credentials -UserName 'CONTOSO\admin'
        $result | Should -Be '$Credsadmin'
    }

    It 'Should sanitize hyphens to underscores' {
        $result = Resolve-Credentials -UserName 'CONTOSO\admin-user'
        $result | Should -Be '$Credsadmin_user'
    }

    It 'Should sanitize dots to underscores' {
        $result = Resolve-Credentials -UserName 'CONTOSO\admin.user'
        $result | Should -Be '$Credsadmin_user'
    }

    It 'Should remove spaces and @ signs' {
        $result = Resolve-Credentials -UserName 'admin @company'
        $result | Should -Be '$Credsadmincompany'
    }

    It 'Should handle a simple username without domain' {
        $result = Resolve-Credentials -UserName 'admin'
        $result | Should -Be '$Credsadmin'
    }
}

Describe 'Add-ReverseDSCUserName' {
    BeforeEach {
        $Script:AllUsers = @()
    }

    It 'Should add a username to the list' {
        Add-ReverseDSCUserName -UserName 'user1@contoso.com'
        $Script:AllUsers | Should -Contain 'user1@contoso.com'
    }

    It 'Should not add duplicate usernames' {
        Add-ReverseDSCUserName -UserName 'user1@contoso.com'
        Add-ReverseDSCUserName -UserName 'user1@contoso.com'
        $Script:AllUsers | Should -HaveCount 1
    }
}

Describe 'Get-ReverseDSCUserNames' {
    BeforeAll {
        $Script:AllUsers = @()
        Add-ReverseDSCUserName -UserName 'user1@contoso.com'
        Add-ReverseDSCUserName -UserName 'user2@contoso.com'
    }

    It 'Should return all added usernames' {
        $result = Get-ReverseDSCUserNames
        $result | Should -HaveCount 2
        $result | Should -Contain 'user1@contoso.com'
        $result | Should -Contain 'user2@contoso.com'
    }
}

Describe 'Clear-ReverseDSCUserNames' {
    BeforeAll {
        Add-ReverseDSCUserName -UserName 'user1@contoso.com'
    }

    It 'Should clear all usernames from the list' {
        Clear-ReverseDSCUserNames
        $Script:AllUsers | Should -HaveCount 0
    }
}

Describe 'Add-ConfigurationDataEntry' {
    BeforeEach {
        Clear-ConfigurationDataContent
    }

    It 'Should add an entry under a new node' {
        Add-ConfigurationDataEntry -Node 'localhost' -Key 'Setting1' -Value 'Value1'
        $result = Get-ConfigurationDataEntry -Node 'localhost' -Key 'Setting1'
        $result.Value | Should -Be 'Value1'
    }

    It 'Should add an entry with a description' {
        Add-ConfigurationDataEntry -Node 'localhost' -Key 'Setting1' -Value 'Value1' -Description 'Test setting'
        $result = Get-ConfigurationDataEntry -Node 'localhost' -Key 'Setting1'
        $result.Value | Should -Be 'Value1'
        $result.Description | Should -Be 'Test setting'
    }

    It 'Should update the value when adding the same key to the same node' {
        Add-ConfigurationDataEntry -Node 'localhost' -Key 'Setting1' -Value 'Value1'
        Add-ConfigurationDataEntry -Node 'localhost' -Key 'Setting1' -Value 'Value2'
        $result = Get-ConfigurationDataEntry -Node 'localhost' -Key 'Setting1'
        $result.Value | Should -Be 'Value2'
    }

    It 'Should support multiple nodes' {
        Add-ConfigurationDataEntry -Node 'Server1' -Key 'Key1' -Value 'A'
        Add-ConfigurationDataEntry -Node 'Server2' -Key 'Key1' -Value 'B'
        (Get-ConfigurationDataEntry -Node 'Server1' -Key 'Key1').Value | Should -Be 'A'
        (Get-ConfigurationDataEntry -Node 'Server2' -Key 'Key1').Value | Should -Be 'B'
    }
}

Describe 'Get-ConfigurationDataEntry' {
    BeforeAll {
        Clear-ConfigurationDataContent
        Add-ConfigurationDataEntry -Node 'localhost' -Key 'TestKey' -Value 'TestValue'
    }

    It 'Should return the entry for a specific node and key' {
        $result = Get-ConfigurationDataEntry -Node 'localhost' -Key 'TestKey'
        $result | Should -Not -BeNullOrEmpty
        $result.Value | Should -Be 'TestValue'
    }

    It 'Should return null when the key does not exist' {
        $result = Get-ConfigurationDataEntry -Node 'localhost' -Key 'NonExistent'
        $result | Should -BeNullOrEmpty
    }
}

Describe 'Clear-ConfigurationDataContent' {
    It 'Should clear all configuration data entries' {
        Add-ConfigurationDataEntry -Node 'localhost' -Key 'TestKey' -Value 'TestValue'
        Clear-ConfigurationDataContent
        $result = Get-ConfigurationDataEntry -Node 'localhost' -Key 'TestKey'
        $result | Should -BeNullOrEmpty
    }
}

Describe 'Get-ConfigurationDataContent' {
    BeforeAll {
        Clear-ConfigurationDataContent
        Add-ConfigurationDataEntry -Node 'localhost' -Key 'ServerName' -Value 'MyServer' -Description 'The server name'
    }

    It 'Should return a string containing the AllNodes section' {
        $result = Get-ConfigurationDataContent
        $result | Should -Match 'AllNodes'
    }

    It 'Should include the node name' {
        $result = Get-ConfigurationDataContent
        $result | Should -Match 'localhost'
    }

    It 'Should include the key and value' {
        $result = Get-ConfigurationDataContent
        $result | Should -Match 'ServerName'
        $result | Should -Match 'MyServer'
    }

    It 'Should include the description as a comment' {
        $result = Get-ConfigurationDataContent
        $result | Should -Match '# The server name'
    }

    It 'Should include NonNodeData section' {
        $result = Get-ConfigurationDataContent
        $result | Should -Match 'NonNodeData'
    }

    It 'Should start with @{ and end with }' {
        $result = Get-ConfigurationDataContent
        $result | Should -Match '^@\{'
        $result | Should -Match '\}$'
    }
}

Describe 'New-ConfigurationDataDocument' {
    BeforeAll {
        Clear-ConfigurationDataContent
        Add-ConfigurationDataEntry -Node 'localhost' -Key 'TestKey' -Value 'TestValue'
        $testPath = Join-Path -Path $TestDrive -ChildPath 'TestConfig.psd1'
    }

    It 'Should create a .psd1 file at the specified path' {
        New-ConfigurationDataDocument -Path $testPath
        Test-Path -Path $testPath | Should -BeTrue
    }

    It 'Should write valid content to the file' {
        New-ConfigurationDataDocument -Path $testPath
        $content = Get-Content -Path $testPath -Raw
        $content | Should -Match 'AllNodes'
        $content | Should -Match 'TestKey'
    }
}

Describe 'ConvertTo-ConfigurationDataString' {
    Context 'When converting a string object' {
        It 'Should wrap the string in quotes with a semicolon' {
            $result = ConvertTo-ConfigurationDataString -PSObject 'TestValue'
            $result | Should -Match '"TestValue"'
        }
    }

    Context 'When converting an array of strings' {
        It 'Should format as a PowerShell array block' {
            $result = ConvertTo-ConfigurationDataString -PSObject @('Item1', 'Item2')
            $result | Should -Match '@\('
            $result | Should -Match 'Item1'
            $result | Should -Match 'Item2'
        }
    }

    Context 'When converting a hashtable' {
        It 'Should format as a PowerShell hashtable block' {
            $hashtable = @{ Name = 'Test' }
            $result = ConvertTo-ConfigurationDataString -PSObject $hashtable
            $result | Should -Match '@\{'
            $result | Should -Match 'Name'
        }
    }
}

Describe 'Convert-DSCStringParamToVariable' {
    Context 'When converting a simple string parameter to a variable' {
        It 'Should remove quotes around the parameter value' {
            $dscBlock = "            ParamName            = `"SomeValue`";`r`n"
            $result = Convert-DSCStringParamToVariable -DSCBlock $dscBlock -ParameterName 'ParamName'
            $result | Should -Not -Match '"SomeValue"'
            $result | Should -Match 'SomeValue'
        }
    }

    Context 'When the parameter name is not found' {
        It 'Should return the original DSCBlock unchanged' {
            $dscBlock = "            OtherParam           = `"Value`";`r`n"
            $result = Convert-DSCStringParamToVariable -DSCBlock $dscBlock -ParameterName 'NonExistent'
            $result | Should -Be $dscBlock
        }
    }
}

Describe 'Get-DSCBlock' {
    BeforeAll {
        # Create a minimal DSC resource module for testing
        $testModulePath = Join-Path -Path $TestDrive -ChildPath 'TestResource.psm1'
        $moduleContent = @'
function Get-TargetResource
{
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter()]
        [System.Boolean]
        $Enabled,

        [Parameter()]
        [System.String[]]
        $Items
    )
}

function Set-TargetResource
{
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter()]
        [System.Boolean]
        $Enabled,

        [Parameter()]
        [System.String[]]
        $Items
    )
}
'@
        Set-Content -Path $testModulePath -Value $moduleContent
    }

    Context 'When generating a DSC block with string parameters' {
        It 'Should produce a properly formatted DSC configuration block' {
            $params = @{
                Name = 'TestResource'
            }
            $result = Get-DSCBlock -ModulePath $testModulePath -Params $params
            $result | Should -Not -BeNullOrEmpty
            $result | Should -Match 'Name'
            $result | Should -Match 'TestResource'
        }
    }

    Context 'When generating a DSC block with boolean parameters' {
        It 'Should format boolean values with $ prefix' {
            $params = @{
                Name    = 'Test'
                Enabled = $true
            }
            $result = Get-DSCBlock -ModulePath $testModulePath -Params $params
            $result | Should -Match '\$True'
        }
    }

    Context 'When generating a DSC block with string array parameters' {
        It 'Should format string arrays with @()' {
            $params = @{
                Name  = 'Test'
                Items = @('Item1', 'Item2')
            }
            $result = Get-DSCBlock -ModulePath $testModulePath -Params $params
            $result | Should -Match '@\('
            $result | Should -Match 'Item1'
            $result | Should -Match 'Item2'
        }
    }

    Context 'When parameters are aligned' {
        It 'Should pad shorter parameter names with spaces for alignment' {
            $params = @{
                Name    = 'Test'
                Enabled = $true
            }
            $result = Get-DSCBlock -ModulePath $testModulePath -Params $params
            # Both parameters should have equal signs, and shorter names should have more spacing
            $result | Should -Match 'Name\s+='
            $result | Should -Match 'Enabled\s+='
        }
    }

    Context 'When _metadata_ properties are present' {
        It 'Should exclude _metadata_ keys from the output but include their values as comments' {
            $params = @{
                Name              = 'Test'
                _metadata_Name    = '# This is a comment'
            }
            $result = Get-DSCBlock -ModulePath $testModulePath -Params $params
            $result | Should -Not -Match '_metadata_'
            $result | Should -Match '# This is a comment'
        }
    }

    Context 'When null values are present' {
        It 'Should exclude parameters with null values' {
            $params = @{
                Name  = 'Test'
                Items = $null
            }
            $result = Get-DSCBlock -ModulePath $testModulePath -Params $params
            # Null params are excluded in the preprocessing step
            $result | Should -Match 'Name'
        }
    }

    Context 'When NoEscape is specified for a parameter' {
        It 'Should not escape the specified parameter values' {
            $params = @{
                Name = '$ConfigName'
            }
            $result = Get-DSCBlock -ModulePath $testModulePath -Params $params -NoEscape @('Name')
            $result | Should -Match '\$ConfigName'
            $result | Should -Not -Match '`\$ConfigName'
        }
    }

    Context 'When hashtable parameters are provided' {
        It 'Should format hashtable values as @{ key = value }' {
            $params = @{
                Name  = 'Test'
                Items = @{ SubKey = 'SubValue' }
            }
            $result = Get-DSCBlock -ModulePath $testModulePath -Params $params
            $result | Should -Match '@\{'
            $result | Should -Match 'SubKey'
        }
    }
}

Describe 'Module Exports' {
    BeforeAll {
        $manifestPath = Join-Path -Path $PSScriptRoot -ChildPath '..\ReverseDSC.psd1'
        $manifest = Test-ModuleManifest -Path $manifestPath -ErrorAction SilentlyContinue
    }

    It 'Should have a valid module manifest' {
        $manifest | Should -Not -BeNullOrEmpty
    }

    It 'Should export expected functions' -ForEach @(
        @{ FunctionName = 'ConvertTo-EscapedDSCString' }
        @{ FunctionName = 'Get-DSCParamType' }
        @{ FunctionName = 'Get-DSCBlock' }
        @{ FunctionName = 'Get-DSCFakeParameters' }
        @{ FunctionName = 'Get-DSCDependsOnBlock' }
        @{ FunctionName = 'Get-Credentials' }
        @{ FunctionName = 'Resolve-Credentials' }
        @{ FunctionName = 'Save-Credentials' }
        @{ FunctionName = 'Test-Credentials' }
        @{ FunctionName = 'Convert-DSCStringParamToVariable' }
        @{ FunctionName = 'Get-ConfigurationDataContent' }
        @{ FunctionName = 'New-ConfigurationDataDocument' }
        @{ FunctionName = 'Add-ConfigurationDataEntry' }
        @{ FunctionName = 'Get-ConfigurationDataEntry' }
        @{ FunctionName = 'Clear-ConfigurationDataContent' }
        @{ FunctionName = 'Add-ReverseDSCUserName' }
    ) {
        Get-Command -Name $FunctionName -Module 'ReverseDSC.Core' -ErrorAction SilentlyContinue |
            Should -Not -BeNullOrEmpty
    }
}

} # InModuleScope

# Cleanup
Remove-Module -Name 'ReverseDSC.Core' -ErrorAction SilentlyContinue
