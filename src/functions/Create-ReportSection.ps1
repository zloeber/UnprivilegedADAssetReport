Function Create-ReportSection
{
    #** This function is specific to this script and does all kinds of bad practice
    #   stuff. Use this function neither to learn from or judge me please. **
    #
    #   That being said, this function pretty much does all the report output
    #   options and layout magic. It depends upon the report layout hash and
    #   $HTMLRendering global variable hash.
    #
    #   This function generally shouldn't need to get changed in any way to customize your
    #   reports.
    #
    # .EXAMPLE
    #    Create-ReportSection -Rpt $ReportSection -Asset $Asset 
    #                         -Section 'Summary' -TableTitle 'System Summary'
    
    [CmdletBinding()]
    param(
        [parameter()]
        $Rpt,
        
        [parameter()]
        [string]$Asset,

        [parameter()]
        [string]$Section,
        
        [parameter()]
        [string]$TableTitle        
    )
    BEGIN
    {
        Add-Type -AssemblyName System.Web
    }
    PROCESS
    {}
    END
    {
        # Get our section type
        $RptSection = $Rpt['Sections'][$Section]
        $SectionType = $RptSection['Type']
        
        switch ($SectionType)
        {
            'Section'     # default to a data section
            {
                Write-Verbose -Message ('Create-ReportSection: {0}: {1}' -f $Asset,$Section)
                $ReportElementSource = @($RptSection['AllData'][$Asset])
                if ((($ReportElementSource.Count -gt 0) -and 
                     ($ReportElementSource[0] -ne $null)) -or 
                     ($RptSection['ShowSectionEvenWithNoData']))
                {
                    $SourceProperties = $RptSection['ReportTypes'][$ReportType]['Properties']
                    
                    #region report section type and layout
                    $TableType = $RptSection['ReportTypes'][$ReportType]['TableType']
                    $ContainerType = $RptSection['ReportTypes'][$ReportType]['ContainerType']

                    switch ($TableType)
                    {
                        'Horizontal' 
                        {
                            $PropertyCount = $SourceProperties.Count
                            $Vertical = $false
                        }
                        'Vertical' {
                            $PropertyCount = 2
                            $Vertical = $true
                        }
                        default {
                            if ((($SourceProperties.Count) -ge $HorizontalThreshold))
                            {
                                $PropertyCount = 2
                                $Vertical = $true
                            }
                            else
                            {
                                $PropertyCount = $SourceProperties.Count
                                $Vertical = $false
                            }
                        }
                    }
                    #endregion report section type and layout
                    
                    $Table = ''
                    If ($PropertyCount -ne 0)
                    {
                        # Create our future HTML table header
                        $SectionLink = '<a href="{0}"></a>' -f $Section
                        $TableHeader = $HTMLRendering['TableTitle'][$HTMLMode] -replace '<0>',$PropertyCount
                        $TableHeader = $SectionLink + ($TableHeader -replace '<1>',$TableTitle)

                        if ($RptSection.ContainsKey('Comment'))
                        {
                            if ($RptSection['Comment'] -ne $false)
                            {
                                $TableComment = $HTMLRendering['TableComment'][$HTMLMode] -replace '<0>',$PropertyCount
                                $TableComment = $TableComment -replace '<1>',$RptSection['Comment']
                                $TableHeader = $TableHeader + $TableComment
                            }
                        }
                        
                        $AllTableElements = @()
                        Foreach ($TableElement in $ReportElementSource)
                        {
                            $AllTableElements += $TableElement | Select $SourceProperties
                        }

                        # If we are creating a vertical table it takes a bit of transformational work
                        if ($Vertical)
                        {
                            $Count = 0
                            foreach ($Element in $AllTableElements)
                            {
                                $Count++
                                $SingleElement = [string]($Element | ConvertTo-PropertyValue | ConvertTo-Html)
                                if ($Rpt['Configuration']['PostProcessingEnabled'])
                                {
                                    # Add class elements for even/odd rows
                                    $SingleElement = Format-HTMLTable $SingleElement -ColorizeMethod 'ByEvenRows' -Attr 'class' -AttrValue 'even' -WholeRow
                                    $SingleElement = Format-HTMLTable $SingleElement -ColorizeMethod 'ByOddRows' -Attr 'class' -AttrValue 'odd' -WholeRow
                                    if ($RptSection.ContainsKey('PostProcessing') -and 
                                       ($RptSection['PostProcessing'].Value -ne $false))
                                    {
                                        $Rpt['Configuration']['PostProcessingEnabled'].Value
                                        $Table = $(Invoke-Command ([scriptblock]::Create($RptSection['PostProcessing'])))
                                    }
                                }
                                $SingleElement = [Regex]::Match($SingleElement, "(?s)(?<=</tr>)(.+)(?=</table>)").Value
                                $Table += $SingleElement 
                                if ($Count -ne $AllTableElements.Count)
                                {
                                    $Table += '<tr class="divide"><td></td><td></td></tr>'
                                }
                            }
                            $Table = '<table class="list">' + $TableHeader + $Table + '</table>'
                            $Table = [System.Web.HttpUtility]::HtmlDecode($Table)
                        }
                        # Otherwise it is a horizontal table
                        else
                        {
                            [string]$Table = $AllTableElements | ConvertTo-Html
                            if ($Rpt['Configuration']['PostProcessingEnabled'])
                            {
                                # Add class elements for even/odd rows
                                $Table = Format-HTMLTable $Table -ColorizeMethod 'ByEvenRows' -Attr 'class' -AttrValue 'even' -WholeRow
                                $Table = Format-HTMLTable $Table -ColorizeMethod 'ByOddRows' -Attr 'class' -AttrValue 'odd' -WholeRow
                                if ($RptSection.ContainsKey('PostProcessing'))
                                
                                {
                                    if ($RptSection.ContainsKey('PostProcessing'))
                                    {
                                        if ($RptSection['PostProcessing'] -ne $false)
                                        {
                                            $Table = $(Invoke-Command ([scriptblock]::Create($RptSection['PostProcessing'])))
                                        }
                                    }
                                }
                            }
                            # This will gank out everything after the first colgroup so we can replace it with our own spanned header
                            $Table = [Regex]::Match($Table, "(?s)(?<=</colgroup>)(.+)(?=</table>)").Value
                            $Table = '<table>' + $TableHeader + $Table + '</table>'
                            $Table = [System.Web.HttpUtility]::HtmlDecode(($Table))
                        }
                    }
                    
                    $Output = $HTMLRendering['SectionContainers'][$HTMLMode][$ContainerType]['Head'] + 
                              $Table + $HTMLRendering['SectionContainers'][$HTMLMode][$ContainerType]['Tail']
                    $Output
                }
            }
            'SectionBreak'
            {
                if ($Rpt['Configuration']['SkipSectionBreaks'] -eq $false)
                {
                    $Output = $HTMLRendering['CustomSections'][$SectionType] -replace '<0>',$TableTitle
                    $Output
                }
            }
        }
    }
}
