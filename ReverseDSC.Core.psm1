$Global:CredsRepo = @()

function Get-DSCParamType
{
<#
.SYNOPSIS
Retrieves the data type of a specific parameter from the associated DSC
resource.

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
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [parameter(Mandatory = $true)]
        [System.String]
        $ModulePath,

        [parameter(Mandatory = $true)]
        [System.String]
        $ParamName
    )

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ModulePath, [ref] $tokens, [ref] $errors)
    $functions = $ast.FindAll( {$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true)

    ForEach($function in $functions){
        if ($function.Name -eq "Set-TargetResource")
        {
            $functionAst = [System.Management.Automation.Language.Parser]::ParseInput($function.Body, [ref] $tokens, [ref] $errors)

            $parameters = $functionAst.FindAll( {$args[0] -is [System.Management.Automation.Language.ParameterAst]}, $true)
            ForEach($parameter in $parameters){
                if ($parameter.Name.Extent.Text -eq $ParamName)
                {
                    $attributes = $parameter.Attributes
                    ForEach($attribute in $attributes){
                        if ($attribute.TypeName.FullName -like "System.*")
                        {
                            return $attribute.TypeName.FullName
                        }
                        elseif ($attribute.TypeName.FullName.ToLower() -eq "microsoft.management.infrastructure.ciminstance")
                        {
                            return "System.Collections.Hashtable"
                        }
                        elseif ($attribute.TypeName.FullName.ToLower() -eq "string")
                        {
                            return "System.String"
                        }
                        elseif ($attribute.TypeName.FullName.ToLower() -eq "boolean")
                        {
                            return "System.Boolean"
                        }
                        elseif ($attribute.TypeName.FullName.ToLower() -eq "bool")
                        {
                            return "System.Boolean"
                        }
                        elseif ($attribute.TypeName.FullName.ToLower() -eq "string[]")
                        {
                            return "System.String[]"
                        }
                        elseif ($attribute.TypeName.FullName.ToLower() -eq "microsoft.management.infrastructure.ciminstance[]")
                        {
                            return "Microsoft.Management.Infrastructure.CimInstance[]"
                        }
                    }
                }
            }
        }
     }
 }

function Get-DSCBlock
{
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

#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $ModulePath,

        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]
        $Params
    )

    $Sorted = $Params.GetEnumerator() | Sort-Object -Property Name
    $NewParams = [Ordered]@{}

    foreach($entry in $Sorted)
    {
        $NewParams.Add($entry.Key, $entry.Value)
    }

    # Figure out what parameter has the longuest name, and get its Length;
    $maxParamNameLength = 0
    foreach ($param in $NewParams.Keys)
    {
        if ($param.Length -gt $maxParamNameLength)
        {
            $maxParamNameLength = $param.Length
        }
    }

    # PSDscRunAsCredential is 20 characters and in most case the longuest.
    if ($maxParamNameLength -lt 20)
    {
        $maxParamNameLength = 20
    }

    $dscBlock = ""
    $NewParams.Keys | ForEach-Object {
        if ($null -ne $NewParams[$_])
        {
            $paramType = $NewParams[$_].GetType().Name
        }
        else
        {
            $paramType = Get-DSCParamType -ModulePath $ModulePath -ParamName "`$$_"
        }

        $value = $null
        if ($paramType -eq "System.String" -or $paramType -eq "String" -or $paramType -eq "Guid" -or $paramType -eq 'TimeSpan')
        {
            if (!$null -eq $NewParams.Item($_))
            {
                $value = "`"" + $NewParams.Item($_).ToString().Replace("`"", "```"") + "`""
            }
            else
            {
                $value = "`"" + $NewParams.Item($_) + "`""
            }
        }
        elseif ($paramType -eq "System.Boolean" -or $paramType -eq "Boolean")
        {
            $value = "`$" + $NewParams.Item($_)
        }
        elseif ($paramType -eq "System.Management.Automation.PSCredential")
        {
            if ($null -ne $NewParams.Item($_))
            {
                if ($NewParams.Item($_).ToString() -like "`$Creds*")
                {
                    $value = $NewParams.Item($_).Replace("-", "_").Replace(".", "_")
                }
                else
                {
                    if ($null -eq $NewParams.Item($_).UserName)
                    {
                        $value = "`$Creds" + ($NewParams.Item($_).Split('\'))[1].Replace("-", "_").Replace(".", "_")
                    }
                    else
                    {
                        if ($NewParams.Item($_).UserName.Contains("@") -and !$NewParams.Item($_).UserName.COntains("\"))
                        {
                            $value = "`$Creds" + ($NewParams.Item($_).UserName.Split('@'))[0]
                        }
                        else
                        {
                            $value = "`$Creds" + ($NewParams.Item($_).UserName.Split('\'))[1].Replace("-", "_").Replace(".", "_")
                        }
                    }
                }
            }
            else
            {
                $value = "Get-Credential -Message " + $_
            }
        }
        elseif ($paramType -eq "System.Collections.Hashtable" -or $paramType -eq "Hashtable")
        {
            $value = "@{"
            $hash = $NewParams.Item($_)
            $hash.Keys | foreach-object {
                try
                {
                    $value += $_ + " = `"" + $hash.Item($_) + "`"; "
                }
                catch
                {
                    $value = $hash
                }
            }
            $value += "}"
        }
        elseif ($paramType -eq "System.String[]" -or $paramType -eq "String[]" -or $paramType -eq "ArrayList" -or $paramType -eq "List``1")
        {
            $hash = $NewParams.Item($_)
            if ($hash -and !$hash.ToString().StartsWith("`$ConfigurationData."))
            {
                $value = "@("
                $hash| ForEach-Object {
                    $value += "`"" + $_ + "`","
                }
                if($value.Length -gt 2)
                {
                    $value = $value.Substring(0,$value.Length -1)
                }
                $value += ")"
            }
            else
            {
                if ($hash)
                {
                    $value = $hash
                }
                else
                {
                    $value = "@()"
                }
            }
        }
        elseif ($paramType -eq "System.UInt32[]")
        {
            $hash = $NewParams.Item($_)
            if ($hash)
            {
                $value = "@("
                $hash| ForEach-Object {
                    $value += $_.ToString() + ","
                }
                if($value.Length -gt 2)
                {
                    $value = $value.Substring(0,$value.Length -1)
                }
                $value += ")"
            }
            else
            {
                if ($hash)
                {
                    $value = $hash
                }
                else
                {
                    $value = "@()"
                }
            }
        }
        elseif ($paramType -eq "Object[]" -or $paramType -eq "Microsoft.Management.Infrastructure.CimInstance[]")
        {
            $array = $hash = $NewParams.Item($_)

            if ($array.Length -gt 0 -and ($array[0].GetType().Name -eq "String" -and $paramType -ne "Microsoft.Management.Infrastructure.CimInstance[]"))
            {
                $value = "@("
                $hash| ForEach-Object {
                    $value += "`"" + $_ + "`","
                }
                if($value.Length -gt 2)
                {
                    $value = $value.Substring(0,$value.Length -1)
                }
                $value += ")"
            }
            else
            {
                $value = "@("
                $array | ForEach-Object{
                    $value += $_
                }
                $value += ")"
            }
        }
        elseif ($paramType -eq "CimInstance")
        {
            $value = $NewParams[$_]
        }
        else
        {
            if ($null -eq $NewParams[$_])
            {
                $value = "`$null"
            }
            else
            {
                if($NewParams[$_].GetType().BaseType.Name -eq "Enum")
                {
                    $value = "`"" + $NewParams.Item($_) + "`""
                }
                else
                {
                    $value = $NewParams.Item($_)
                }
            }
        }

        # Determine the number of additional spaces we need to add before the '=' to make sure the values are all aligned. This number
        # is obtained by substracting the length of the current parameter's name to the maximum length found.
        $numberOfAdditionalSpaces = $maxParamNameLength - $_.Length
        $additionalSpaces = ""
        for ($i = 0; $i -lt $numberOfAdditionalSpaces; $i++)
        {
            $additionalSpaces += " "
        }
        $dscBlock += "            " + $_  + $additionalSpaces + " = " + $value + ";`r`n"
    }

    return $dscBlock
}

function Get-DSCFakeParameters
{
<#
.SYNOPSIS
Generates a hashtable containing all the properties exposed by the specified
DSC resource but with fake values.

.DESCRIPTION
This function scans the specified resources, create a hashtable with all the
properties it exposes and generates fake values for each property based on
the Data Type assigned to it.

.PARAMETER ModulePath
Full file path to the .psm1 module we are looking to get an instance of.
In most cases this will be the full path to the .psm1 file of the DSC resource.

#>
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
    $functions = $ast.FindAll( {$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true)

    $functions | ForEach-Object{

        if ($_.Name -eq "Get-TargetResource")
        {
            $functionAst = [System.Management.Automation.Language.Parser]::ParseInput($_.Body, [ref] $tokens, [ref] $errors)

            $parameters = $functionAst.FindAll( {$args[0] -is [System.Management.Automation.Language.ParameterAst]}, $true)
            $parameters | ForEach-Object{
                $paramName = $_.Name.Extent.Text
                $attributes = $_.Attributes
                $found = $false

                <# Loop once to figure out if there is a validate Set to use. #>
                $attributes | ForEach-Object{
                    if ($_.TypeName.FullName -eq "ValidateSet")
                    {
                        $params.Add($paramName.Replace("`$", ""), $_.PositionalArguments[0].ToString().Replace("`"", "").Replace("'",""))
                        $found = $true
                    }
                    elseif ($_.TypeName.FullName -eq "ValidateRange") {
                        $params.Add($paramName.Replace("`$", ""), $_.PositionalArguments[0].ToString())
                        $found = $true
                    }
                }
                $attributes | ForEach-Object{
                    if (!$found)
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
                            $params.Add($paramName.Replace("`$", ""),[string]@("1","2"))
                            $found = $true
                        }
                    }
                }
            }
        }
     }
     return $params
}

function Get-DSCDependsOnBlock
{
<#
.SYNOPSIS
Generates a string that represents the DependsOn clause based on the received
list of dependencies.

.DESCRIPTION
This function receives an array of string that represents the list of DSC
resource dependencies for the current DSC block and generates a string
that represents the associated DependsOn DSC string.

.PARAMETER DependsOnItems
Array of string values that represent the list of depdencies for the
current DSC block. Object in the array are expected to be in the form of:
[<DSCResourceName>]<InstanceName>.

#>
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
    $dependsOnClause = $dependsOnClause.Substring(0, $dependsOnClause.Length -1)
    $dependsOnClause += ");"
    return $dependsOnClause
}

<# Region Helper Methods #>
function Get-Credentials
{
<#
.SYNOPSIS
Returns the full username of (<domain>\<username>) of the specified user
if it is already stroed in our credentials hashtable.

.DESCRIPTION
This function checks in the hashtable that stores all the required
credentials (service account, etc.) for our configuration and
returns the fully formatted username.

.PARAMETER UserName
Name of the user we wish to check to see if it is already stored in our
credentials hashtable.

#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $UserName
    )
    
    if ($Global:CredsRepo.Contains($UserName.ToLower()))
    {
        return $UserName.ToLower()
    }
    return $null
}

function Resolve-Credentials
{
<#
.SYNOPSIS
Returns a string representing the name of the PSCredential variable
associated with the specific username.

.DESCRIPTION
This function takes in a specified user name and returns what the standardized
variable name for that user should be inside of our extracted DSC configuration.
Credentials variables will always be named $Creds<username> as a standard for
ReverseDSC. This function makes sure that the variable name doesn't contain
character that are invalid in variable names bu might be valid in Usernames.

.PARAMETER UserName
Name of the user we wish to get the associated variable name from.

#>
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
        return "`$Creds" + $userNameParts[1].Replace("-","_").Replace(".", "_").Replace(" ", "").Replace("@","")
    }
    return "`$Creds" + $UserName.Replace("-","_").Replace(".", "_").Replace(" ", "").Replace("@","")
}

function Save-Credentials
{
<#
.SYNOPSIS
Adds the specified username to our central list of required credentials.

.DESCRIPTION
This function checks to see if the specified user is already stored in our
central required credentials list, and if not simply adds it to it.

.PARAMETER UserName
Username to add to the central list of required credentials.

#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $UserName
    )
    if (!$Global:CredsRepo.Contains($UserName.ToLower()))
    {
        $Global:CredsRepo += $UserName.ToLower()
    }
}

function Test-Credentials
{
<#
.SYNOPSIS
Checks to see if the specified username if already in our central list of
required credentials.

.DESCRIPTION
This function checks the central list of required credentials to see if the
specified user is already part of it. If it finds it, it returns $true,
otherwise it returns false.

.PARAMETER UserName
Username to check for existence in the central list of required users.

#>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $UserName
    )
    if ($Global:CredsRepo.Contains($UserName.ToLower()))
    {
        return $true
    }
    return $false
}

function Convert-DSCStringParamToVariable
{
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
array of CIM instances or not. We need to differentiate by explicitely
passing in this parameter because to the function a CIMArray is nothing
but a System.Object[] and will threat it as it. CIMArray differ in that
we should not have commas in between items it contains.

#>
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
        $IsCIMArray = $false
    )
    # There is a possibility that a parameter's value will contain the name of the property
    # or that another property will contain a subset of it.
    $startPosition = $DSCBlock.IndexOf($ParameterName)

    # If quotes appear before an equal sign, when starting from the assumed start position,
    # then the start position is invalid, searhc for another instance of the Parameter;
    $startPosition = -1
    do
    {
        $startPosition = $DSCBlock.IndexOf($ParameterName, $startPosition + 1)
        $testValidStartPositionEqual = $DSCBlock.IndexOf("=", $startPosition)
        $testValidStartPositionQuotes = $DSCBlock.IndexOf("`"", $startPosition)
    } while ($testValidStartPositionEqual -gt $testValidStartPositionQuotes -and 
             $startPosition -ne -1)

    $endOfLinePosition = $DSCBlock.IndexOf(";`r`n", $startPosition)

    if ($endOfLinePosition -eq -1)
    {
        $endOfLinePosition = $DSCBlock.Length
    }
    $startPosition = $DSCBlock.IndexOf("`"", $startPosition)

    while ($startPosition -ge 0 -and $startPosition -lt $endOfLinePosition)
    {
        if ($endOfLinePosition -gt $startPosition)
        {
            if ($startPosition -ge 0)
            {
                $endPosition = $DSCBlock.IndexOf("`"", $startPosition + 1)
                if ($endPosition -lt 0)
                {
                    $endPosition = $DSCBlock.IndexOf("'", $startPosition + 1)
                }

                if ($endPosition -ge 0 -and $endPosition -le $endofLinePosition)
                {
                    $removeBeginQuotes = $true
                    $removeEndQuotes = $true
                    $NewStartPosition = $startPosition
                    if ($IsCIMArray)
                    {
                        $previousEqualSignPosition = $DSCBlock.IndexOf("=", $startPosition -2)

                        # If we have  a CIMArray, and the current quote we are looking at
                        # is exactly 2 positions before it, we skip remove it because it
                        # actually is the quotes surrounding a value of an entry of the
                        # CIMArray. If it was the principal quotes we were looking at removing
                        # the previous equal sign would be further before due to CIMArray being
                        # declared as ' = @("MSFT_....';
                        if (($previousEqualSignPosition - $startPosition -2) -lt 0)
                        {
                            $removeBeginQuotes = $false
                        }
                        
                        $previousEqualSignPosition = $DSCBlock.IndexOf("=", $endPosition -2)
                        $nextNewLinePosition = $DSCBLock.IndexOf("`r`n", $endPosition + 1)
                        if (($previousEqualSignPosition - $endPosition -2) -lt 0 -or
                             $nextNewLinePosition -eq ($endPosition +1))
                        {
                            $removeEndQuotes = $false
                            $newStartPosition = $DSCBlock.IndexOf("`r`n", $endPosition)
                        }
                    }
                    
                    if ($removeBeginQuotes)
                    {
                        $DSCBlock = $DSCBlock.Remove($startPosition, 1)
                    }
                    if ($removeEndQuotes)
                    {
                        $DSCBlock = $DSCBlock.Remove($endPosition-1, 1)
                    }

                    $startPosition = $newStartPosition
                }
                else
                {
                    $startPosition = -1
                }
            }
        }
        $startPosition = $DSCBlock.IndexOf("`"", $startPosition)
    }

    if($IsCIMArray)
    {
        $DSCBlock = $DSCBlock.Replace("}`r`n,", "`}`r`n")
    }
    return $DSCBlock
}

<# Region ConfigurationData Methods #>
$ConfigurationDataContent = @{}

function Add-ConfigurationDataEntry
{
<#
.SYNOPSIS
Adds a property to the resulting ConfigurationData file from the
extract.

.DESCRIPTION
This function helps build the hashtable that will eventually result
in the ConfigurationData .psd1 file generated by the extraction of
the configuration. It allows you to speficy what section to add it
to inside the hashtable, and allows you to speficy a description for
each one. These description will eventually become comments that
will appear on top of the property in the ConfigurationData file.

.PARAMETER Node
Specifies the node entry under which we want to add this parameter
under. You can also specify NonNodeData names to have the property
added under custom non-node specific section.

.PARAMETER Key
The name of the parameter to add.

.PARAMETER Value
The value of the parameter to add.

.PARAMETER Description
Description of the parameter to add. This will ultimately appear in
the generated ConfigurationData .psd1 file as a comment appearing on
top of the parameter.

#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $Node,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Key,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Value, 

        [Parameter()]
        [System.String]
        $Description
    )
    if ($null -eq $ConfigurationDataContent[$Node])
    {
        $ConfigurationDataContent.Add($Node, @{})
        $ConfigurationDataContent[$Node].Add("Entries", @{})
    }
    if (!$ConfigurationDataContent[$Node].Entries.ContainsKey($Key))
    {
        $ConfigurationDataContent[$Node].Entries.Add($Key, @{Value = $Value; Description = $Description})
    }
}

function Get-ConfigurationDataEntry
{
<#
.SYNOPSIS
Retrieves the value of a given property in the specified node/section
from the hashtable that is being dynamically built.

.DESCRIPTION
This function will return the value of the specified parameter from the
hash table being dynamically built and which will ultimately become the
content of the ConfigurationData .psd1 file being generated.

.PARAMETER Node
The name of the node or section in the Hashtable we want to look for
the key in.

.PARAMETER Key
The name of the parameter to retrieve the value from.

#>
    [CmdletBinding()]
    [OutputType([System.String])]
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
        foreach ($Node in $ConfigurationDataContent.Keys)
        {
            if ($ConfigurationDataContent[$Node].Entries.ContainsKey($Key))
            {
                return $ConfigurationDataContent[$Node].Entries[$Key]
            }
        }
    }
    else
    {
        if ($ConfigurationDataContent[$Node].Entries.ContainsKey($Key))
        {
            return $ConfigurationDataContent[$Node].Entries[$Key]
        }
    }
}

function Get-ConfigurationDataContent
{
<#
.SYNOPSIS
Retrieves the entire content of the ConfigurationData file being
dynamically generated.

.DESCRIPTION
This function will return the content of the dynamically built
hashtable for the ConfigurationData content as a formatted string.

#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param()

    $psd1Content = "@{`r`n"
    $psd1Content += "    AllNodes = @(`r`n"
    foreach ($node in $ConfigurationDataContent.Keys.Where{$_.ToLower() -ne "nonnodedata"})
    {
        $psd1Content += "        @{`r`n"
        $psd1Content += "            NodeName                    = `"" + $node + "`"`r`n"
        $psd1Content += "            PSDscAllowPlainTextPassword = `$true;`r`n"
        $psd1Content += "            PSDscAllowDomainUser        = `$true;`r`n"
        $psd1Content += "            # CertificateFile           = `"\\<location>\<file>.cer`";`r`n"
        $psd1Content += "            # Thumbprint                = `xxxxxxxxxxxxxxxxxxxxxxx`r`n`r`n"
        $psd1Content += "            #region Parameters`r`n"
        $keyValuePair = $ConfigurationDataContent[$node].Entries
        foreach ($key in $keyValuePair.Keys)
        {
            if ($null -ne $keyValuePair[$key].Description)
            {
                $psd1Content += "            # " + $keyValuePair[$key].Description +  "`r`n"
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
        $psd1Content = $psd1Content.Remove($psd1Content.Length-3, 1)
    }

    $psd1Content += "    )`r`n"
    $psd1Content += "    NonNodeData = @(`r`n"
    foreach ($node in $ConfigurationDataContent.Keys.Where{$_.ToLower() -eq "nonnodedata"})
    {
        $psd1Content += "        @{`r`n"
        $keyValuePair = $ConfigurationDataContent[$node].Entries
        foreach ($key in $keyValuePair.Keys)
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
                    $newValue = $newValue.Substring(0,$newValue.Length -1)
                    $newValue += ")"
                    $value = $newValue
                }
            
                if ($null -ne $keyValuePair[$key].Description)
                {
                    $psd1Content += "            # " + $keyValuePair[$key].Description +  "`r`n"
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
            catch {
                Write-Host "Warning: Could not obtain value for key $key" -ForegroundColor Yellow
            }
        }
        $psd1Content += "        }`r`n"
    }
    if($psd1Content.EndsWith(",`r`n"))
    {
        $psd1Content = $psd1Content.Remove($psd1Content.Length-3, 1)
    }
    $psd1Content += "    )`r`n"
    $psd1Content += "}"
    return $psd1Content
}

function New-ConfigurationDataDocument
{
<#
.SYNOPSIS
Generates a new ConfigurationData .psd1 file.

.DESCRIPTION
This function will create the ConfigurationData .psd1 file and store
the content of the converted hashtable in it.

.PARAMETER Path
Full file path of the the resulting file will be located.

#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $Path
    )
    Get-ConfigurationDataContent | Out-File -FilePath $Path
}

function ConvertTo-ConfigurationDataString
{
<#
.SYNOPSIS
Converts items from the content of the dynamic hashtable to be used as
the content of the ConfigurationData .psd1 file into their proper string
representation.

.DESCRIPTION
This function will loop through all items inside the dynamic hashtable
used for the resulting ConfigurationData .psd1 file's content and
converts each one to the proper string representation based on their
data type.

.PARAMETER PSObject
The hashtable object we are building and which is to be used to drive
the content of the ConfigurationData .psd1 file.

#>
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
                $configDataContent = $configDataContent.Remove($configDataContent.Length-3, 1)
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

<# Region User based Methods #>
$Global:AllUsers = @()

function Add-ReverseDSCUserName
{
<#
.SYNOPSIS
Adds the provided username to the list of required users for the
destination environment.

.DESCRIPTION
ReverseDSC allows you to keep track of all user credentials encountered
during various stages of the extraction process. By keeping a central list
of all users account required by the source environment we can easily
generate a script that will automatically create new user place holders
in a destination environment's Active Directory. This function checks
to see if the specified user was already encountered, and if not adds it
to the central list of all required users.

.PARAMETER UserName
Name of the user to add to the central list of required users.

#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $UserName
    )
    if (!$Global:AllUsers.Contains($UserName))
    {
        $Global:AllUsers += $UserName
    }
}
