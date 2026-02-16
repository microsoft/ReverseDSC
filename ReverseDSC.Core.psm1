$Script:CredsRepo = @()

<#
.SYNOPSIS
    Escapes a string for use in DSC configuration blocks.

.DESCRIPTION
    Applies standard escaping rules for backticks, European localized
    quotation marks (U+201E, U+201C, U+201D), and double quotes.
    Optionally preserves PowerShell variable expressions ($...).

.PARAMETER InputString
    The raw string value to escape.

.PARAMETER AllowVariables
    When specified, dollar signs ($) are not escaped, allowing
    PowerShell variable expansion in the resulting string.
#>
function ConvertTo-EscapedDSCString
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [System.String]
        $InputString,

        [Parameter()]
        [switch]
        $AllowVariables
    )

    if ([System.String]::IsNullOrEmpty($InputString))
    {
        return [System.String]::Empty
    }

    $result = $InputString.Replace('`', '``')
    if (-not $AllowVariables)
    {
        $result = $result.Replace('$', '`$')
    }
    # Escape European localized quotation marks (U+201E „, U+201C “, U+201D ”)
    $result = $result.Replace("$([char]0x201E)", "``$([char]0x201E)")
    $result = $result.Replace("$([char]0x201C)", "``$([char]0x201C)")
    $result = $result.Replace("$([char]0x201D)", "``$([char]0x201D)")
    $result = $result.Replace("`"", "```"")
    return $result
}

<#
.SYNOPSIS
    Retrieves the data type of a specific parameter from the associated DSC resource.

.DESCRIPTION
    This function scans the specified module (or in this case DSC resource),
    checks for the specified parameter inside the .schema.mof file associated
    with that module and properly assesses and returns the Data Type assigned
    to the parameter.

.PARAMETER ModulePath
    Full file path to the .psm1 module we are looking for the property inside of.
    In most cases this will be the full path to the .psm1 file of the DSC resource.

.PARAMETER ParamName
    Name of the parameter in the module we want to determine the Data Type for.
#>
function Get-DSCParamType
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $ModulePath,

        [Parameter(Mandatory = $true)]
        [System.String]
        $ParamName
    )

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ModulePath, [ref] $tokens, [ref] $errors)
    $functions = $ast.FindAll( { $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)

    foreach ($function in $functions)
    {
        if ($function.Name -eq "Set-TargetResource")
        {
            $functionAst = [System.Management.Automation.Language.Parser]::ParseInput($function.Body, [ref] $tokens, [ref] $errors)

            $parameters = $functionAst.FindAll( { $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)
            foreach ($parameter in $parameters)
            {
                if ($parameter.Name.Extent.Text -eq $ParamName)
                {
                    $attributes = $parameter.Attributes
                    foreach ($attribute in $attributes)
                    {
                        if ($attribute.TypeName.FullName -like "System.*")
                        {
                            return $attribute.TypeName.FullName
                        }
                        elseif ($attribute.TypeName.FullName -eq "Microsoft.Management.Infrastructure.CimInstance")
                        {
                            return "System.Collections.Hashtable"
                        }
                        elseif ($attribute.TypeName.FullName -eq "string")
                        {
                            return "System.String"
                        }
                        elseif ($attribute.TypeName.FullName -eq "boolean")
                        {
                            return "System.Boolean"
                        }
                        elseif ($attribute.TypeName.FullName -eq "bool")
                        {
                            return "System.Boolean"
                        }
                        elseif ($attribute.TypeName.FullName -eq "string[]")
                        {
                            return "System.String[]"
                        }
                        elseif ($attribute.TypeName.FullName -eq "Microsoft.Management.Infrastructure.CimInstance[]")
                        {
                            return "Microsoft.Management.Infrastructure.CimInstance[]"
                        }
                    }
                }
            }
        }
    }
}

<#
.SYNOPSIS
    Converts a string parameter value to its DSC representation.

.PARAMETER Value
    The string value to convert.

.PARAMETER NoEscape
    If true, the string will not be escaped.

.PARAMETER AllowVariables
    If true, PowerShell variables ($...) are preserved in the string.
#>
function ConvertTo-DSCStringValue
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter()]
        [System.String]
        $Value,

        [Parameter()]
        [System.Boolean]
        $NoEscape = $false,

        [Parameter()]
        [System.Boolean]
        $AllowVariables = $false
    )

    if ($null -eq $Value)
    {
        return '""'
    }

    if ($NoEscape)
    {
        return $Value
    }

    $escapedString = ConvertTo-EscapedDSCString -InputString $Value -AllowVariables:$AllowVariables
    return "`"$escapedString`""
}

<#
.SYNOPSIS
    Converts a boolean parameter value to its DSC representation.

.PARAMETER Value
    The boolean value to convert.
#>
function ConvertTo-DSCBooleanValue
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Boolean]
        $Value
    )

    return "`$$Value"
}

<#
.SYNOPSIS
    Converts a PSCredential parameter value to its DSC representation.

.PARAMETER Value
    The PSCredential value to convert.

.PARAMETER ParameterName
    The name of the parameter (used as fallback).
#>
function ConvertTo-DSCCredentialValue
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Value,

        [Parameter(Mandatory = $true)]
        [System.String]
        $ParameterName
    )

    if ($null -eq $Value)
    {
        return "Get-Credential -Message $ParameterName"
    }

    $credString = $Value.ToString()
    if ($credString -like "`$Creds*")
    {
        return $credString.Replace("-", "_").Replace(".", "_")
    }

    $userName = $Value.UserName
    if ($null -eq $userName)
    {
        $userName = ($credString.Split('\'))[1]
    }

    if ($userName.Contains("@") -and -not $userName.Contains("\"))
    {
        $cleanName = ($userName.Split('@'))[0]
    }
    else
    {
        $cleanName = ($userName.Split('\'))[-1]
    }

    $cleanName = $cleanName.Replace("-", "_").Replace(".", "_").Replace(" ", "").Replace("@", "")
    return "`$Creds$cleanName"
}

<#
.SYNOPSIS
    Converts a hashtable parameter value to its DSC representation.

.PARAMETER Value
    The hashtable to convert.
#>
function ConvertTo-DSCHashtableValue
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]
        $Value
    )

    $result = "@{"
    foreach ($key in $Value.Keys)
    {
        try
        {
            $result += "$key = `"$($Value[$key])`"; "
        }
        catch
        {
            return $Value
        }
    }
    $result += "}"
    return $result
}

<#
.SYNOPSIS
    Converts a string array parameter value to its DSC representation.

.PARAMETER Value
    The array to convert.

.PARAMETER NoEscape
    If true, array elements will not be escaped.

.PARAMETER AllowVariables
    If true, PowerShell variables in strings are preserved.
#>
function ConvertTo-DSCStringArrayValue
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter()]
        [System.Object[]]
        $Value,

        [Parameter()]
        [System.Boolean]
        $NoEscape = $false,

        [Parameter()]
        [System.Boolean]
        $AllowVariables = $false
    )

    if ($null -eq $Value -or $Value.Count -eq 0)
    {
        return "@()"
    }

    $result = "@("
    foreach ($item in $Value)
    {
        if ($null -ne $item)
        {
            if ($NoEscape)
            {
                $innerValue = $item
            }
            else
            {
                $innerValue = ConvertTo-EscapedDSCString -InputString $item -AllowVariables:$AllowVariables
            }
            $result += "`"$innerValue`","
        }
    }

    if ($result.Length -gt 2 -and $result.EndsWith(","))
    {
        $result = $result.Substring(0, $result.Length - 1)
    }
    $result += ")"
    return $result
}

<#
.SYNOPSIS
    Converts an integer array parameter value to its DSC representation.

.PARAMETER Value
    The integer array to convert.
#>
function ConvertTo-DSCIntegerArrayValue
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter()]
        [System.Object[]]
        $Value
    )

    if ($null -eq $Value -or $Value.Count -eq 0)
    {
        return "@()"
    }

    return "@($($Value -join ','))"
}

<#
.SYNOPSIS
    Converts an object array parameter value to its DSC representation.

.PARAMETER Value
    The array to convert.

.PARAMETER NoEscape
    If true, string elements will not be escaped.

.PARAMETER AllowVariables
    If true, PowerShell variables in strings are preserved.
#>
function ConvertTo-DSCObjectArrayValue
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter()]
        [System.Object[]]
        $Value,

        [Parameter()]
        [System.Boolean]
        $NoEscape = $false,

        [Parameter()]
        [System.Boolean]
        $AllowVariables = $false
    )

    if ($null -eq $Value -or $Value.Count -eq 0)
    {
        return "@()"
    }

    # Handle string arrays
    if ($Value[0].GetType().Name -eq "String")
    {
        $result = "@("
        foreach ($item in $Value)
        {
            if ($NoEscape)
            {
                $result += $item
            }
            else
            {
                $escapedString = ConvertTo-EscapedDSCString -InputString $item -AllowVariables:$AllowVariables
                $result += "`"$escapedString`","
            }
        }

        # Remove the trailing comma if it exists
        if ($result.Length -gt 2 -and $result.EndsWith(","))
        {
            $result = $result.Substring(0, $result.Length - 1)
        }
        $result += ")"
        return $result
    }

    # Handle hashtable arrays
    if ($Value[0].GetType().Name -eq "Hashtable")
    {
        $result = "@("
        foreach ($hashtable in $Value)
        {
            $result += "@{"
            foreach ($pair in $hashtable.GetEnumerator())
            {
                if ($pair.Value -is [System.Array])
                {
                    $str = "$($pair.Key)=@('$($pair.Value -join "', '")')"
                }
                else
                {
                    if ($null -eq $pair.Value)
                    {
                        $str = "$($pair.Key)=`$null"
                    }
                    else
                    {
                        $str = "$($pair.Key)='$($pair.Value)'"
                    }
                }
                $result += "$str; "
            }

            # Remove the trailing semicolon and space if they exist
            if ($result.Length -gt 2 -and $result.EndsWith("; "))
            {
                $result = $result.Substring(0, $result.Length - 2)
            }
            $result += "}"
        }
        $result += ")"
        return $result
    }

    # Default handling for other object arrays
    $result = "@("
    foreach ($item in $Value)
    {
        $result += $item
    }
    $result += ")"
    return $result
}

<#
.SYNOPSIS
    Generate the DSC string representing the resource's instance.
.DESCRIPTION
    This function is really the core of ReverseDSC. It takes in an array of
    parameters and returns the DSC string that represents the given instance
    of the specified resource.
.PARAMETER ModulePath
    Full file path to the .psm1 module we are looking to get an instance of.
    In most cases this will be the full path to the .psm1 file of the DSC resource.
.PARAMETER Params
    Hashtable that contains the list of Key properties and their values.
.PARAMETER NoEscape
    Array of string values that represent the list of parameters that should
    not be escaped when generating the DSC string.
#>
function Get-DSCBlock
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $ModulePath,

        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]
        $Params,

        [Parameter()]
        [System.String[]]
        $NoEscape,

        [Parameter()]
        [switch]
        $AllowVariablesInStrings
    )

    # Sort the params by name(key), exclude _metadata_* properties (coming from DSCParser)
    $Sorted = $Params.GetEnumerator() | Where-Object Name -NotLike '_metadata_*' | Sort-Object -Property Name
    $NewParams = [ordered]@{}

    foreach ($entry in $Sorted)
    {
        if ($null -ne $entry.Value)
        {
            $NewParams.Add($entry.Key, $entry.Value)
        }
    }

    # Figure out what parameter has the longest name, and get its Length;
    $maxParamNameLength = 0
    foreach ($param in $NewParams.Keys)
    {
        if ($param.Length -gt $maxParamNameLength)
        {
            $maxParamNameLength = $param.Length
        }
    }

    # PSDscRunAsCredential is 20 characters and in most case the longest.
    if ($maxParamNameLength -lt 20)
    {
        $maxParamNameLength = 20
    }

    $dscBlock = [System.Text.StringBuilder]::new()
    $NewParams.Keys | ForEach-Object {
        $paramName = $_
        $paramValue = $NewParams[$paramName]

        if ($null -ne $paramValue)
        {
            $paramType = $paramValue.GetType().Name
        }
        else
        {
            $paramType = Get-DSCParamType -ModulePath $ModulePath -ParamName "`$$paramName"
        }

        $isNoEscape = $NoEscape -contains $paramName
        $value = $null

        # Dispatch to type-specific converter
        switch -Regex ($paramType)
        {
            '^(System\.String|String|Guid|TimeSpan|DateTime)$'
            {
                $value = ConvertTo-DSCStringValue -Value $paramValue -NoEscape $isNoEscape -AllowVariables $AllowVariablesInStrings
            }
            '^(System\.Boolean|Boolean)$'
            {
                $value = ConvertTo-DSCBooleanValue -Value $paramValue
            }
            '^System\.Management\.Automation\.PSCredential$'
            {
                $value = ConvertTo-DSCCredentialValue -Value $paramValue -ParameterName $paramName
            }
            '^(System\.Collections\.Hashtable|Hashtable)$'
            {
                $value = ConvertTo-DSCHashtableValue -Value $paramValue
            }
            '^(System\.String\[\]|String\[\]|ArrayList|List``1)$'
            {
                if ($paramValue.ToString().StartsWith("`$ConfigurationData."))
                {
                    $value = $paramValue
                }
                else
                {
                    $value = ConvertTo-DSCStringArrayValue -Value $paramValue -NoEscape $isNoEscape -AllowVariables $AllowVariablesInStrings
                }
            }
            'Int.*\[\]'
            {
                $value = ConvertTo-DSCIntegerArrayValue -Value $paramValue
            }
            '^(Object\[\]|Microsoft\.Management\.Infrastructure\.CimInstance\[\])$'
            {
                if ($paramType -ne "Microsoft.Management.Infrastructure.CimInstance[]" -and
                    $paramValue.Length -gt 0 -and $paramValue[0].GetType().Name -eq "String")
                {
                    $value = ConvertTo-DSCObjectArrayValue -Value $paramValue -NoEscape $isNoEscape -AllowVariables $AllowVariablesInStrings
                }
                else
                {
                    $value = ConvertTo-DSCObjectArrayValue -Value $paramValue -NoEscape $isNoEscape -AllowVariables $AllowVariablesInStrings
                }
            }
            '^CimInstance$'
            {
                $value = $paramValue
            }
            default
            {
                if ($null -eq $paramValue)
                {
                    $value = "`$null"
                }
                elseif ($paramValue.GetType().BaseType.Name -eq "Enum")
                {
                    $value = "`"$paramValue`""
                }
                else
                {
                    $value = "$paramValue"
                }
            }
        }

        # Determine the number of additional spaces we need to add before the '=' to make sure the values are all aligned. This number
        # is obtained by subtracting the length of the current parameter's name from the maximum length found.
        $numberOfAdditionalSpaces = $maxParamNameLength - $paramName.Length
        $additionalSpaces = " " * $numberOfAdditionalSpaces

        # Check for comment/metadata and insert it back here
        $PropertyMetadataKeyName = "_metadata_$paramName"
        if ($Params.ContainsKey($PropertyMetadataKeyName))
        {
            $CommentValue = ' ' + $Params[$PropertyMetadataKeyName]
        }
        else
        {
            $CommentValue = ''
        }
        [void]$dscBlock.Append("            $paramName$additionalSpaces = $value;$CommentValue`r`n")
    }

    return $dscBlock.ToString()
}

<#
.SYNOPSIS
    Generates a hashtable containing all the properties exposed by the
    specified DSC resource but with fake values.

.DESCRIPTION
    This function scans the specified resource, creates a hashtable with all
    the properties it exposes and generates fake values for each property
    based on the Data Type assigned to it.

.PARAMETER ModulePath
    Full file path to the .psm1 module we are looking to get an instance of.
    In most cases this will be the full path to the .psm1 file of the DSC resource.
#>
function Get-DSCFakeParameters
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $ModulePath
    )

    $params = @{}

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ModulePath, [ref] $tokens, [ref] $errors)
    $functions = $ast.FindAll( { $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)

    $functions | ForEach-Object {

        if ($_.Name -eq "Get-TargetResource")
        {
            $functionAst = [System.Management.Automation.Language.Parser]::ParseInput($_.Body, [ref] $tokens, [ref] $errors)

            $parameters = $functionAst.FindAll( { $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)
            $parameters | ForEach-Object {
                $paramName = $_.Name.Extent.Text
                $attributes = $_.Attributes
                $found = $false

                <# Loop once to figure out if there is a validate Set to use. #>
                $attributes | ForEach-Object {
                    if ($_.TypeName.FullName -eq "ValidateSet")
                    {
                        $params.Add($paramName.Replace("`$", ""), $_.PositionalArguments[0].ToString().Replace("`"", "").Replace("'", ""))
                        $found = $true
                    }
                    elseif ($_.TypeName.FullName -eq "ValidateRange")
                    {
                        $params.Add($paramName.Replace("`$", ""), $_.PositionalArguments[0].ToString())
                        $found = $true
                    }
                }
                $attributes | ForEach-Object {
                    if (-not $found)
                    {
                        if ($_.TypeName.FullName -eq "System.String" -or $_.TypeName.FullName -eq "String")
                        {
                            $params.Add($paramName.Replace("`$", ""), "*")
                            $found = $true
                        }
                        elseif ($_.TypeName.FullName -eq "System.UInt32" -or $_.TypeName.FullName -eq "Int32")
                        {
                            $params.Add($paramName.Replace("`$", ""), 0)
                            $found = $true
                        }
                        elseif ($_.TypeName.FullName -eq "System.Management.Automation.PSCredential")
                        {
                            $params.Add($paramName.Replace("`$", ""), $null)
                            $found = $true
                        }
                        elseif ($_.TypeName.FullName -eq "System.Management.Automation.Boolean" -or $_.TypeName.FullName -eq "System.Boolean" -or $_.TypeName.FullName -eq "Boolean")
                        {
                            $params.Add($paramName.Replace("`$", ""), $true)
                            $found = $true
                        }
                        elseif ($_.TypeName.FullName -eq "System.String[]" -or $_.TypeName.FullName -eq "String[]")
                        {
                            $params.Add($paramName.Replace("`$", ""), [string]@("1", "2"))
                            $found = $true
                        }
                    }
                }
            }
        }
    }
    return $params
}

<#
.SYNOPSIS
    Generates a string that represents the DependsOn clause based on the
    received list of dependencies.

.DESCRIPTION
    This function receives an array of strings that represents the list of DSC
    resource dependencies for the current DSC block and generates a string
    that represents the associated DependsOn DSC string.

.PARAMETER DependsOnItems
    Array of string values that represent the list of dependencies for the
    current DSC block. Objects in the array are expected to be in the form of:
    [<DSCResourceName>]<InstanceName>.
#>
function Get-DSCDependsOnBlock
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Object[]]
        $DependsOnItems
    )
    $dependsOnClause = "@("
    foreach ($clause in $DependsOnItems)
    {
        $dependsOnClause += "`"" + $clause + "`","
    }
    $dependsOnClause = $dependsOnClause.Substring(0, $dependsOnClause.Length - 1)
    $dependsOnClause += ");"
    return $dependsOnClause
}

<#
.SYNOPSIS
    Returns the full username (<domain>\<username>) of the specified user
    if it is already stored in the credentials hashtable.

.DESCRIPTION
    This function checks in the hashtable that stores all the required
    credentials (service account, etc.) for the configuration and
    returns the fully formatted username.

.PARAMETER UserName
    Name of the user we wish to check to see if it is already stored in our
    credentials hashtable.
#>
function Get-Credentials
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $UserName
    )

    if ($Script:CredsRepo.Contains($UserName.ToLower()))
    {
        return $UserName.ToLower()
    }
    return $null
}

<#
.SYNOPSIS
    Returns a string representing the name of the PSCredential variable
    associated with the specified username.

.DESCRIPTION
    This function takes in a specified user name and returns what the
    standardized variable name for that user should be inside of our
    extracted DSC configuration. Credential variables will always be named
    $Creds<username> as a standard for ReverseDSC. This function makes sure
    that the variable name does not contain characters that are invalid in
    variable names but might be valid in usernames.

.PARAMETER UserName
    Name of the user we wish to get the associated variable name from.
#>
function Resolve-Credentials
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $UserName
    )
    $userNameParts = $UserName.ToLower().Split('\')
    if ($userNameParts.Length -gt 1)
    {
        return "`$Creds" + $userNameParts[1].Replace("-", "_").Replace(".", "_").Replace(" ", "").Replace("@", "")
    }
    return "`$Creds" + $UserName.Replace("-", "_").Replace(".", "_").Replace(" ", "").Replace("@", "")
}

<#
.SYNOPSIS
    Adds the specified username to our central list of required credentials.

.DESCRIPTION
    This function checks to see if the specified user is already stored in our
    central required credentials list, and if not simply adds it to it.

.PARAMETER UserName
    Username to add to the central list of required credentials.
#>
function Save-Credentials
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $UserName
    )
    if (-not $Script:CredsRepo.Contains($UserName.ToLower()))
    {
        $Script:CredsRepo += $UserName.ToLower()
    }
}

<#
.SYNOPSIS
    Checks to see if the specified username is already in our central list of
    required credentials.

.DESCRIPTION
    This function checks the central list of required credentials to see if
    the specified user is already part of it. If it finds it, it returns
    $true, otherwise it returns $false.

.PARAMETER UserName
    Username to check for existence in the central list of required users.
#>
function Test-Credentials
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $UserName
    )
    if ($Script:CredsRepo.Contains($UserName.ToLower()))
    {
        return $true
    }
    return $false
}

<#
.SYNOPSIS
    Removes quotes around a parameter in the resulting DSC config,
    effectively converting it to a variable instead of a string value.

.DESCRIPTION
    This function will scan the content of the current DSC block for the
    resource, find the specified parameter and remove quotes around its
    value so that it becomes a variable instead of a string value.

.PARAMETER DSCBlock
    The string representation of the current DSC resource instance we
    are extracting along with all of its parameters and values.

.PARAMETER ParameterName
    The name of the parameter we wish to convert the value as a variable
    instead of a string value for.

.PARAMETER IsCIMArray
    Represents whether or not the parameter to convert to a variable is an
    array of CIM instances or not. We need to differentiate by explicitly
    passing in this parameter because to the function a CIMArray is nothing
    but a System.Object[] and will treat it as such. CIMArrays differ in
    that we should not have commas in between items they contain.

.PARAMETER IsCIMObject
    Represents whether or not the parameter to convert to a variable is a
    CIM instance or not. We need to differentiate by explicitly passing
    in this parameter because to the function a CIMObject is nothing
    but a String object and will treat it as such. However it has escaped
    double quotes, which need to be handled properly.
#>
function Convert-DSCStringParamToVariable
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $DSCBlock,

        [Parameter(Mandatory = $true)]
        [System.String]
        $ParameterName,

        [Parameter()]
        [System.Boolean]
        $IsCIMArray = $false,

        [Parameter()]
        [System.Boolean]
        $IsCIMObject = $false
    )

    # If quotes appear before an equal sign, when starting from the assumed start position,
    # then the start position is invalid, search for another instance of the Parameter;
    $startPosition = -1
    do
    {
        $startPosition = $DSCBlock.IndexOf(' ' + $ParameterName + ' ', $startPosition + 1)
        # If the ParameterName is not found, $startPosition is still -1, and .IndexOf($string, $startPosition) does not work
        if ($startPosition -ne -1)
        {
            $testValidStartPositionEqual = $DSCBlock.IndexOf("=", $startPosition)
            $testValidStartPositionQuotes = $DSCBlock.IndexOf("`"", $startPosition)
        }
    } while ($testValidStartPositionEqual -gt $testValidStartPositionQuotes -and
        $startPosition -ne -1)

    # If $ParameterName was not found i.e. $startPosition is still -1, skip this section as well.
    # We just want the original DSCBlock to be returned.
    if ($startPosition -ne -1) {
        $endOfLinePosition = $DSCBlock.IndexOf(";`r`n", $startPosition)

        if ($endOfLinePosition -eq -1)
        {
            $endOfLinePosition = $DSCBlock.Length
        }
        $startPosition = $DSCBlock.IndexOf("`"", $startPosition)
    }

    while ($startPosition -ge 0 -and $startPosition -lt $endOfLinePosition)
    {
        $endOfLinePosition = $DSCBlock.IndexOf(";`r`n", $startPosition)

        if ($endOfLinePosition -eq -1)
        {
            $endOfLinePosition = $DSCBlock.Length
        }
        if ($endOfLinePosition -gt $startPosition)
        {
            if ($startPosition -ge 0)
            {
                $endPosition = $DSCBlock.IndexOf("`"", $startPosition + 1)
                 <#
                    When the parameter is a CIM array, it may contain parameter with double quotes
                    We need to ensure that endPosition does not correspond to such parameter
                    by checking if the second character before " is =
                    Additionally, there might be other values in the DSC block, e.g. from xml,
                    which contain other properties like <?xml version="1.0"?>, where we do
                    not want to remove the quotes as well.
                #>
                if ($IsCIMArray -or $IsCIMObject)
                {
                    while ($endPosition -gt 1 -and `
                        ($DSCBlock.Substring($endPosition -3,4) -eq '= `"' -or `
                         $DSCBlock.Substring($endPosition -2,3) -eq '=`"'))
                    {
                        #This retrieve the endquote that we skip
                        $endPosition = $DSCBlock.IndexOf("`"", $endPosition + 1)
                        #This retrieve the next quote
                        $endPosition = $DSCBlock.IndexOf("`"", $endPosition + 1)
                    }
                }
                if ($endPosition -lt 0)
                {
                    $endPosition = $DSCBlock.IndexOf("'", $startPosition + 1)
                }

                if ($endPosition -ge 0 -and $endPosition -le $endofLinePosition)
                {
                    $DSCBlock = $DSCBlock.Remove($startPosition, 1)
                    $DSCBlock = $DSCBlock.Remove($endPosition - 1, 1)
                }
                else
                {
                    $startPosition = -1
                }
            }
        }
        $startPosition = $DSCBlock.IndexOf("`"", $startPosition)
        <#
            When the parameter is a CIM array, it may contain parameter with double quotes
            We need to ensure that startPosition does not correspond to such parameter
            by checking if the third character before " is =
            Additionally, there might be other values in the DSC block, e.g. from xml,
            which contain other properties like <?xml version="1.0"?>, where we do
            not want to remove the quotes as well.
        #>
        if ($IsCIMArray -or $IsCIMObject)
        {
            while ($startPosition -gt 1 -and `
                ($DSCBlock.Substring($startPosition -3,4) -eq '= `"' -or `
                 $DSCBlock.Substring($startPosition -2,3) -eq '=`"'))
            {
                #This retrieve the endquote that we skip
                $startPosition = $DSCBlock.IndexOf("`"", $startPosition + 1)
                #This retrieve the next quote
                $startPosition = $DSCBlock.IndexOf("`"", $startPosition + 1)
            }
        }
    }

    if ($IsCIMArray -or $IsCIMObject)
    {
        $DSCBlock = $DSCBlock.Replace("},`r`n", "`}`r`n")
        $DSCBlock = $DSCBlock -replace "`r`n\s*[,;]`r`n", "`r`n" # replace "<crlf>[<whitespace>][,;]<crlf>" with "<crlf>"

        # There are cases where the closing ')' of a CIMInstance array still has leading quotes.
        # This ensures we clean those out.
        $indexOfProperty = $DSCBlock.IndexOf($ParameterName)
        if ($indexOfProperty -ge 0)
        {
            $indexOfEndOfLine = $DSCBlock.IndexOf(";`r`n", $indexOfProperty)
            if ($indexOfEndOfLine -gt 0 -and $indexOfEndOfLine -gt $indexOfProperty)
            {
                $propertyString = $DSCBlock.Substring($indexOfProperty, $indexOfEndOfLine - $indexOfProperty + 1)
                if ($propertyString.EndsWith("}`");"))
                {
                    $fixedPropertyString = $propertyString.Replace("}`");", "}`r`n            );")
                    $DSCBlock = $DSCBLock.Replace($propertyString, $fixedPropertyString)
                }
            }

            # Correcting escaped double quotes to non-escaped double quotes
            $indexOfEndOfLine = $DSCBlock.IndexOf(";`r`n", $indexOfProperty)
            if ($indexOfEndOfLine -gt 0 -and $indexOfEndOfLine -gt $indexOfProperty)
            {
                $propertyString = $DSCBlock.Substring($indexOfProperty, $indexOfEndOfLine - $indexOfProperty + 1)
                if ($propertyString.Contains('`"'))
                {
                    $fixedPropertyString = $propertyString.Replace('`"', '"')
                    $DSCBlock = $DSCBLock.Replace($propertyString, $fixedPropertyString)
                }
            }
        }
        #$DSCBlock = $DSCBLock.Replace('}");', "}`r`n            )")
    }
    return $DSCBlock
}

$Script:ConfigurationDataContent = @{}

<#
.SYNOPSIS
    Adds a property to the resulting ConfigurationData file from the extract.

.DESCRIPTION
    This function helps build the hashtable that will eventually result
    in the ConfigurationData .psd1 file generated by the extraction of
    the configuration. It allows you to specify what section to add it
    to inside the hashtable, and allows you to specify a description for
    each one. These descriptions will eventually become comments that
    will appear on top of the property in the ConfigurationData file.

.PARAMETER Node
    Specifies the node entry under which we want to add this parameter.
    You can also specify NonNodeData names to have the property added
    under custom non-node specific sections.

.PARAMETER Key
    The name of the parameter to add.

.PARAMETER Value
    The value of the parameter to add.

.PARAMETER Description
    Description of the parameter to add. This will ultimately appear in
    the generated ConfigurationData .psd1 file as a comment appearing on
    top of the parameter.
#>
function Add-ConfigurationDataEntry
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $Node,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Key,

        [Parameter(Mandatory = $true)]
        [System.Object]
        $Value,

        [Parameter()]
        [System.String]
        $Description
    )

    if ($null -eq $Script:ConfigurationDataContent[$Node])
    {
        $Script:ConfigurationDataContent.Add($Node, @{})
        $Script:ConfigurationDataContent[$Node].Add("Entries", [ordered]@{})
    }

    $Script:ConfigurationDataContent[$Node].Entries[$Key] = @{ Value = $Value; Description = $Description }
}

<#
.SYNOPSIS
    Retrieves the value of a given property in the specified node/section
    from the hashtable that is being dynamically built.

.DESCRIPTION
    This function will return the value of the specified parameter from the
    hashtable being dynamically built and which will ultimately become the
    content of the ConfigurationData .psd1 file being generated.

.PARAMETER Node
    The name of the node or section in the hashtable we want to look for
    the key in.

.PARAMETER Key
    The name of the parameter to retrieve the value from.
#>
function Get-ConfigurationDataEntry
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $Node,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Key
    )
    <# If node is null, then search in all nodes and return first result found. #>
    if ($null -eq $Node)
    {
        foreach ($Node in $Script:ConfigurationDataContent.Keys)
        {
            if ($Script:ConfigurationDataContent[$Node].Entries.Contains($Key))
            {
                return $Script:ConfigurationDataContent[$Node].Entries[$Key]
            }
        }
    }
    else
    {
        if ($Script:ConfigurationDataContent.ContainsKey($Node) -and $Script:ConfigurationDataContent[$Node].Entries.Contains($Key))
        {
            return $Script:ConfigurationDataContent[$Node].Entries[$Key]
        }
    }
}

<#
.SYNOPSIS
    Clears the content of the hashtable that is being dynamically built for the ConfigurationData .psd1 file.

.DESCRIPTION
    This function will clear the content of the hashtable that is being built
    for the ConfigurationData .psd1 file, effectively resetting it to an empty
    state.
#>
function Clear-ConfigurationDataContent
{
    [CmdletBinding()]
    param()
    $Script:ConfigurationDataContent = @{}
}

<#
.SYNOPSIS
    Retrieves the entire content of the ConfigurationData file being
    dynamically generated.

.DESCRIPTION
    This function will return the content of the dynamically built
    hashtable for the ConfigurationData content as a formatted string.
#>
function Get-ConfigurationDataContent
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param()

    $psd1Content = "@{`r`n"
    $psd1Content += "    AllNodes = @(`r`n"
    foreach ($node in $Script:ConfigurationDataContent.Keys.Where{ $_ -ne "NonNodeData" })
    {
        $psd1Content += "        @{`r`n"
        $psd1Content += "            NodeName                    = `"" + $node + "`"`r`n"
        $psd1Content += "            PSDscAllowPlainTextPassword = `$true;`r`n"
        $psd1Content += "            PSDscAllowDomainUser        = `$true;`r`n"
        $psd1Content += "            #region Parameters`r`n"
        $keyValuePair = $Script:ConfigurationDataContent[$node].Entries
        foreach ($key in $keyValuePair.Keys | Sort-Object)
        {
            if ($null -ne $keyValuePair[$key].Description)
            {
                $psd1Content += "            # " + $keyValuePair[$key].Description + "`r`n"
            }
            if ($keyValuePair[$key].Value.ToString().StartsWith("@(") -or $keyValuePair[$key].Value.ToString().StartsWith("`$"))
            {
                $psd1Content += "            " + $key + " = " + $keyValuePair[$key].Value + "`r`n`r`n"
            }
            elseif ($keyValuePair[$key].Value.GetType().FullName -eq "System.Object[]")
            {
                $psd1Content += "            " + $key + " = " + (ConvertTo-ConfigurationDataString $keyValuePair[$key].Value)
            }
            else
            {
                $psd1Content += "            " + $key + " = `"" + $keyValuePair[$key].Value + "`"`r`n`r`n"
            }
        }

        $psd1Content += "        },`r`n"
    }

    if ($psd1Content.EndsWith(",`r`n"))
    {
        $psd1Content = $psd1Content.Remove($psd1Content.Length - 3, 1)
    }

    $psd1Content += "    )`r`n"
    $psd1Content += "    NonNodeData = @(`r`n"
    foreach ($node in $Script:ConfigurationDataContent.Keys.Where{ $_ -eq "NonNodeData" })
    {
        $psd1Content += "        @{`r`n"
        $keyValuePair = $Script:ConfigurationDataContent[$node].Entries
        foreach ($key in $keyValuePair.Keys | Sort-Object)
        {
            try
            {
                $value = $keyValuePair[$key].Value
                $valType = $value.GetType().FullName

                if ($valType -eq "System.Object[]")
                {
                    $newValue = "@("
                    foreach ($item in $value)
                    {
                        $newValue += "`"" + $item + "`","
                    }
                    $newValue = $newValue.Substring(0, $newValue.Length - 1)
                    $newValue += ")"
                    $value = $newValue
                }

                if ($null -ne $keyValuePair[$key].Description)
                {
                    $psd1Content += "            # " + $keyValuePair[$key].Description + "`r`n"
                }
                if ($value.StartsWith("@(") -or $value.StartsWith("`$"))
                {
                    $psd1Content += "            " + $key + " = " + $value + "`r`n`r`n"
                }
                else
                {
                    $psd1Content += "            " + $key + " = `"" + $value + "`"`r`n`r`n"
                }
            }
            catch
            {
                Write-Host "Warning: Could not obtain value for key $key" -ForegroundColor Yellow
            }
        }
        $psd1Content += "        }`r`n"
    }
    if ($psd1Content.EndsWith(",`r`n"))
    {
        $psd1Content = $psd1Content.Remove($psd1Content.Length - 3, 1)
    }
    $psd1Content += "    )`r`n"
    $psd1Content += "}"
    return $psd1Content
}

<#
.SYNOPSIS
    Generates a new ConfigurationData .psd1 file.

.DESCRIPTION
    This function will create the ConfigurationData .psd1 file and store
    the content of the converted hashtable in it.

.PARAMETER Path
    Full file path where the resulting file will be located.
#>
function New-ConfigurationDataDocument
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $Path
    )
    Get-ConfigurationDataContent | Out-File -FilePath $Path
}

<#
.SYNOPSIS
    Converts items from the content of the dynamic hashtable to their
    proper string representation for the ConfigurationData .psd1 file.

.DESCRIPTION
    This function will loop through all items inside the dynamic hashtable
    used for the resulting ConfigurationData .psd1 file's content and
    converts each one to the proper string representation based on their
    data type.

.PARAMETER PSObject
    The hashtable object we are building and which is to be used to drive
    the content of the ConfigurationData .psd1 file.
#>
function ConvertTo-ConfigurationDataString
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSObject]
        $PSObject
    )
    $configDataContent = ""
    $objectType = $PSObject.GetType().FullName
    switch ($objectType)
    {
        "System.String"
        {
            $configDataContent += "`"" + $PSObject + "`";`r`n"
        }
        "System.Object[]"
        {
            $configDataContent += "            @(`r`n"
            foreach ($entry in $PSObject)
            {
                $configDataContent += ConvertTo-ConfigurationDataString $entry
            }
            if ($configDataContent.EndsWith(",`r`n"))
            {
                $configDataContent = $configDataContent.Remove($configDataContent.Length - 3, 1)
            }
            $configDataContent += "            )`r`n"
        }

        "System.Collections.Hashtable"
        {
            $configDataContent += "            @{`r`n"
            foreach ($key in $PSObject.Keys)
            {
                $configDataContent += "                " + $key + " = "
                $configDataContent += ConvertTo-ConfigurationDataString $PSObject[$key]
            }
            $configDataContent += "            },`r`n"
        }
    }
    return $configDataContent
}

$Script:AllUsers = @()

<#
.SYNOPSIS
    Adds the provided username to the list of required users for the
    destination environment.

.DESCRIPTION
    ReverseDSC allows you to keep track of all user credentials encountered
    during various stages of the extraction process. By keeping a central
    list of all user accounts required by the source environment we can
    easily generate a script that will automatically create new user
    placeholders in a destination environment's Active Directory. This
    function checks to see if the specified user was already encountered,
    and if not adds it to the central list of all required users.

.PARAMETER UserName
    Name of the user to add to the central list of required users.
#>
function Add-ReverseDSCUserName
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $UserName
    )
    if (-not $Script:AllUsers.Contains($UserName))
    {
        $Script:AllUsers += $UserName
    }
}

<#
.SYNOPSIS
    Retrieves the list of all user accounts required by the source
    environment.

.DESCRIPTION
    This function returns the list of all user accounts that were
    encountered during the extraction process and which are required for
    the configuration to work in the destination environment. This list is
    built by calling the Add-ReverseDSCUserName function every time a new
    user account is encountered during the extraction.
#>
function Get-ReverseDSCUserNames
{
    [CmdletBinding()]
    [OutputType([System.String[]])]
    param()
    return $Script:AllUsers
}

<#
.SYNOPSIS
    Clears the list of all user accounts required by the source environment.

.DESCRIPTION
    This function clears the list of all user accounts that were
    encountered during the extraction process and which are required for
    the configuration to work in the destination environment. This can be
    useful to call at the beginning of an extraction to ensure that you are
    starting with a clean slate in terms of required user accounts.
#>
function Clear-ReverseDSCUserNames
{
    [CmdletBinding()]
    param()
    $Script:AllUsers = @()
}
