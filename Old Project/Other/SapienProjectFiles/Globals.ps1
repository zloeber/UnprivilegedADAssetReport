# Global Variables
$OutputToGrid = $false
$FunctionName = 'YourFunction'
$FunctionPath = '.\New-ADAssetReport.ps1'
$FunctionIsExternal = $true

$FunctionParamTypes = @{}
$FunctionParamTypes.Add("combobox_ReportFormat","combobox")
$FunctionParamTypes.Add("combobox_ReportType","combobox")
$FunctionParamTypes.Add("checkbox_ExportAllUsers","checkbox")
$FunctionParamTypes.Add("checkbox_ExportPrivilegedUsers","checkbox")
$FunctionParamTypes.Add("checkbox_ExportGraphvizDefinitionFiles","checkbox")
$FunctionParamTypes.Add("checkbox_SaveData","checkbox")
$FunctionParamTypes.Add("checkbox_LoadData","checkbox")
$FunctionParamTypes.Add("textbox_DataFile","string")
$FunctionParamTypes.Add("checkbox_Verbose","checkbox")

$MandatoryParams = @{}
$MandatoryParams.Add("ReportFormat",$False)
$MandatoryParams.Add("ReportType",$False)
$MandatoryParams.Add("ExportAllUsers",$False)
$MandatoryParams.Add("ExportPrivilegedUsers",$False)
$MandatoryParams.Add("ExportGraphvizDefinitionFiles",$False)
$MandatoryParams.Add("SaveData",$False)
$MandatoryParams.Add("LoadData",$False)
$MandatoryParams.Add("DataFile",$False)
$MandatoryParams.Add("Verbose",$False)

# Global functions
Function Convert-SplatToString ($Splat)
{
    BEGIN
    {
        Function Escape-PowershellString ([string]$InputString)
        {
            $replacements = @{
                '!' = '`!' 
                '"' = '`"'
                '$' = '`$'
                '%' = '`%'
                '*' = '`*'
                "'" = "`'"
                ' ' = '` '
                '#' = '`#'
                '@' = '`@'
                '.' = '`.'
                '=' = '`='
                ',' = '`,'
            }

            # Join all (escaped) keys from the hashtable into one regular expression.
            [regex]$r = @($replacements.Keys | foreach { [regex]::Escape( $_ ) }) -join '|'

            $matchEval = { param( [Text.RegularExpressions.Match]$matchInfo )
              # Return replacement value for each matched value.
              $matchedValue = $matchInfo.Groups[0].Value
              $replacements[$matchedValue]
            }

            $InputString | Foreach { $r.Replace( $_, $matchEval ) }
        }
    }
    PROCESS
    {
    }
    END
    {
        $ResultSplat = ''
        Foreach ($SplatName in $Splat.Keys)
        {
            switch ((($Splat[$SplatName]).GetType()).Name) {
            	'Boolean' {
            		if ($Splat[$SplatName] -eq $true)
                    {
                        $SplatVal = '$true'
                    }
                    else
                    {
                        $SplatVal = '$false'
                    }
            		break
            	}
            	'String' {
            		$SplatVal = '"' + $(Escape-PowershellString $Splat[$SplatName]) + '"'
                    break
            	}
            	default {
                    $SplatVal = $Splat[$SplatName]
            		break
            	}
            }
            $ResultSplat = $ResultSplat + '-' + $SplatName + ':' + $SplatVal + ' '
        }
        $ResultSplat
    }
}