Function Load-AssetDataFile ($FileToLoad)
{
    $ReportStructure = Import-Clixml -Path $FileToLoad
    # Export/Import XMLCLI isn't going to deal with our embedded scriptblocks (named expressions)
    # so we manually convert them back to scriptblocks like the rockstars we are...
    Foreach ($Key in $ReportStructure['Sections'].Keys) 
    {
        if ($ReportStructure['Sections'][$Key]['Type'] -eq 'Section')  # if not a section break
        {
            Foreach ($ReportTypeKey in $ReportStructure['Sections'][$Key]['ReportTypes'].Keys)
            {
                $ReportStructure['Sections'][$Key]['ReportTypes'][$ReportTypeKey]['Properties'] | 
                    ForEach {
                        $_['e'] = [Scriptblock]::Create($_['e'])
                    }
            }
        }
    }
    Return $ReportStructure
}