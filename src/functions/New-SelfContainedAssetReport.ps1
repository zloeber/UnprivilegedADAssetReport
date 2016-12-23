Function New-SelfContainedAssetReport
{
    <#
    .SYNOPSIS
        Generates a new asset report from gathered data.
    .DESCRIPTION
        Generates a new asset report from gathered data. The information 
        gathering routine generates the output root elements.
    .PARAMETER ReportContainer
        The custom report hash vaiable structure you plan to report upon.
    .PARAMETER DontGatherData
        If your report container already has all the data from a prior run and
        you are just creating a different kind of report with the same data, enable this switch
    .PARAMETER ReportType
        The report type.
    .PARAMETER HTMLMode
        The HTML rendering type (DynamicGrid or EmailFriendly).
    .PARAMETER ExportToExcel
        Export an excel document.
    .PARAMETER EmailRelay
        Email server to relay report through.
    .PARAMETER EmailSender
        Email sender.
    .PARAMETER EmailRecipient
        Email recipient.
    .PARAMETER EmailSubject
        Email subject.
    .PARAMETER SendMail
        Send email of resulting report?
    .PARAMETER ForceAnonymous
        Force email to be sent anonymously?
    .PARAMETER SaveReport
        Save the report?
    .PARAMETER SaveAsPDF
        Save the report as a PDF. If the PDF library is not available the default format, HTML, will be used instead.
    .PARAMETER OutputMethod
        If saving the report, will it be one big report or individual reports?
    .PARAMETER ReportName
        If saving the report, what do you want to call it? This is only used if one big report is being generated.
    .PARAMETER ReportNamePrefix
        Prepend an optional prefix to the report name?
    .PARAMETER ReportLocation
        If saving multiple reports, where will they be saved?
    .EXAMPLE
        New-SelfContainedAssetReport -ReportContainer $ADForestReport -ExportToExcel `
            -SaveReport `
            -OutputMethod 'IndividualReport' `
            -HTMLMode 'DynamicGrid'

        Description:
        ------------------
        Create a forest active directory report.
    .NOTES
        Version    : 1.0.0 10/15/2013
                     - First release

        Author     : Zachary Loeber

        Disclaimer : This script is provided AS IS without warranty of any kind. I 
                     disclaim all implied warranties including, without limitation,
                     any implied warranties of merchantability or of fitness for a 
                     particular purpose. The entire risk arising out of the use or
                     performance of the sample scripts and documentation remains
                     with you. In no event shall I be liable for any damages 
                     whatsoever (including, without limitation, damages for loss of 
                     business profits, business interruption, loss of business 
                     information, or other pecuniary loss) arising out of the use of or 
                     inability to use the script or documentation. 

        Copyright  : I believe in sharing knowledge, so this script and its use is 
                     subject to : http://creativecommons.org/licenses/by-sa/3.0/
    .LINK
        http://www.the-little-things.net/
    .LINK
        http://nl.linkedin.com/in/zloeber

    #>

    #region Parameters
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory=$true, HelpMessage='The custom report hash variable structure you plan to report upon')]
        $ReportContainer,
        
        [Parameter(HelpMessage='Do not gather data, this assumes $Reportcontainer has been pre-populated.')]
        [switch]$DontGatherData,
        
        [Parameter( HelpMessage='The report type')]
        [string]$ReportType = '',
        
        [Parameter( HelpMessage='The HTML rendering type (DynamicGrid or EmailFriendly)')]
        [ValidateSet('DynamicGrid','EmailFriendly')]
        [string]$HTMLMode = 'DynamicGrid',
        
        [Parameter( HelpMessage='Export an excel document as part of the output')]
        [switch]$ExportToExcel,
        
        [Parameter( HelpMessage='Skip html/pdf generation, only produce an excel report (if switch is enabled)')]
        [switch]$NoReport,
        
        [Parameter( HelpMessage='Email server to relay report through')]
        [string]$EmailRelay = '.',
        
        [Parameter( HelpMessage='Email sender')]
        [string]$EmailSender='systemreport@localhost',
     
        [Parameter( HelpMessage='Email recipient')]
        [string]$EmailRecipient='default@yourdomain.com',
        
        [Parameter( HelpMessage='Email subject')]
        [string]$EmailSubject='System Report',
        
        [Parameter( HelpMessage='Send email of resulting report?')]
        [switch]$SendMail,
        
        [Parameter( HelpMessage="Force email to be sent anonymously?")]
        [switch]$ForceAnonymous,

        [Parameter( HelpMessage='Save the report?')]
        [switch]$SaveReport,
        
        [Parameter( HelpMessage='Save the data gathered for later processing?')]
        [switch]$SaveData,
        
        [Parameter( HelpMessage='Save the data gathered for later processing?')]
        [string]$SaveDataFile='DataFile.xml',
        
        [Parameter( HelpMessage='Skip information gathering?')]
        [switch]$SkipInformationGathering,
        
        [Parameter( HelpMessage='Save the report as a PDF. If the PDF library is not available the default format, HTML, will be used instead.')]
        [switch]$SaveAsPDF,

        [Parameter( HelpMessage='Zip up the report(s)?')]
        [switch]$ZipReport,
       
        [Parameter( HelpMessage='How to process report output?')]
        [ValidateSet('OneBigReport','IndividualReport','NoReport')]
        [string]$OutputMethod='OneBigReport',
        
        [Parameter( HelpMessage='If saving the report, what do you want to call it?')]
        [string]$ReportName='Report.html',
        
        [Parameter( HelpMessage='Prepend an optional prefix to the report name?')]
        [string]$ReportNamePrefix='',
        
        [Parameter( HelpMessage='If saving multiple reports, where will they be saved?')]
        [string]$ReportLocation='.'
    )
    #endregion Parameters
    Begin {
        # Use this to keep a splat of our CmdletBinding options
        $VerboseDebug=@{}
        If ($PSBoundParameters.ContainsKey('Verbose')) {
            If ($PSBoundParameters.Verbose -eq $true) {
                $VerboseDebug.Verbose = $true
            }
            else {
                $VerboseDebug.Verbose = $false
            }
        }
        If ($PSBoundParameters.ContainsKey('Debug')) {
            If ($PSBoundParameters.Debug -eq $true) {
                $VerboseDebug.Debug = $true 
            } 
            else {
                $VerboseDebug.Debug = $false
            }
        }

        $ReportOutputSplat = @{
            'SaveAsPDF' = $SaveAsPDF
        }
        
        # Some basic initialization
        $AssetReports = ''
        $FinishedReportPaths = @()
        
        if (($ReportType -eq '') -or ($ReportContainer['Configuration']['ReportTypes'] -notcontains $ReportType)) {
            $ReportType = $ReportContainer['Configuration']['ReportTypes'][0]
            Write-Verbose "New-SelfContainedAssetReport: ReportType set to $ReportType"
        }
        # There must be a more elegant way to do this hash sorting but this also allows
        # us to pull a list of only the sections which are defined and need to be generated.
        $SortedReports = @()
        Foreach ($Key in $ReportContainer['Sections'].Keys) {
            Write-Verbose "New-SelfContainedAssetReport: Processing Section $Key"
            if ($ReportContainer['Sections'][$Key]['ReportTypes'].ContainsKey($ReportType)) {
                if ($ReportContainer['Sections'][$Key]['Enabled'] -and 
                    ($ReportContainer['Sections'][$Key]['ReportTypes'][$ReportType] -ne $false)) {
                    $_SortedReportProp = @{
                        'Section' = $Key
                         'Order' = $ReportContainer['Sections'][$Key]['Order']
                    }
                    $SortedReports += New-Object -Type PSObject -Property $_SortedReportProp
                }
            }
        }
        $SortedReports = $SortedReports | Sort-Object Order
    }
    Process {}
    End {
        if ($SkipInformationGathering){
            Write-Verbose "Skipping information gathering..."
            $AssetNames = @($ReportContainer['Configuration']['Assets'])
        }
        else {
            # Information Gathering, Your custom script block must return the 
            #   array of strings (keys) which consist of the Root elements of your
            #   desired reports.
            Write-Verbose -Message ('New-SelfContainedAssetReport: Invoking information gathering script...')
            $AssetNames = @(Invoke-Command ([scriptblock]::Create($ReportContainer['Configuration']['PreProcessing'])))
        }

        if ($AssetNames.Count -ge 1) {
            Write-Verbose "Found at least one result to process..."
            if ($SaveData){
                Write-Verbose "Saving the data to $ReportNamePrefix$SaveDataFile"
                $ReportContainer | Export-CliXml -Path ($ReportNamePrefix + $SaveDataFile)
            }
            # if we are to export all data to excel, then we do so per section
            #   then per Asset
            if ($ExportToExcel) {
                Write-Verbose -Message ('New-SelfContainedAssetReport: Exporting to excel...')
                # First make sure we have data to export, this shlould also weed out non-data sections meant for html
                #  (like section breaks and such)
                $ProcessExcelReport = $false
                foreach ($ReportSection in $SortedReports){
                    if ($ReportContainer['Sections'][$ReportSection.Section]['AllData'].Count -gt 0) {
                        $ProcessExcelReport = $true
                    }
                }

                #region Excel
                if ($ProcessExcelReport) {
                    # Create the excel workbook
                    try {
                        $Excel = New-Object -ComObject Excel.Application -ErrorAction Stop
                        $ExcelExists = $True
                        $Excel.visible = $True
                        #Start-Sleep -s 1
                        $Workbook = $Excel.Workbooks.Add()
                        $Excel.DisplayAlerts = $false
                    }
                    catch {
                        Write-Warning ('Issues opening excel: {0}' -f $_.Exception.Message)
                        $ExcelExists = $False
                    }
                    if ($ExcelExists)
                    {
                        # going through every section, but in reverse so it shows up in the correct
                        #  sheet in excel. 
                        $SortedExcelReports = $SortedReports | Sort-Object Order -Descending
                        Foreach ($ReportSection in $SortedExcelReports)
                        {
                            $SectionData = $ReportContainer['Sections'][$ReportSection.Section]['AllData']
                            $SectionProperties = $ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType]['Properties']
                            
                            # Gather all the asset information in the section (remember that each asset may
                            #  be pointing to an array of psobjects)
                            $TransformedSectionData = @()                        
                            foreach ($asset in $SectionData.Keys)
                            {
                                # Get all of our calculated properties, then add in the asset name
                                $TempProperties = $SectionData[$asset] | Select $SectionProperties
                                $TransformedSectionData += ($TempProperties | Select @{n='AssetName';e={$asset}},*)
                            }
                            if (($TransformedSectionData.Count -gt 0) -and ($TransformedSectionData -ne $null))
                            {
                                $temparray1 = $TransformedSectionData | ConvertTo-MultiArray
                                if ($temparray1 -ne $null)
                                {    
                                    $temparray = $temparray1.Value
                                    $starta = [int][char]'a' - 1
                                    
                                    if ($temparray.GetLength(1) -gt 26) 
                                    {
                                        $col = [char]([int][math]::Floor($temparray.GetLength(1)/26) + $starta) + [char](($temparray.GetLength(1)%26) + $Starta)
                                    } 
                                    else 
                                    {
                                        $col = [char]($temparray.GetLength(1) + $starta)
                                    }
                                    
                                    Start-Sleep -s 1
                                    $xlCellValue = 1
                                    $xlEqual = 3
                                    $BadColor = 13551615    #Light Red
                                    $BadText = -16383844    #Dark Red
                                    $GoodColor = 13561798    #Light Green
                                    $GoodText = -16752384    #Dark Green
                                    $Worksheet = $Workbook.Sheets.Add()
                                    $Worksheet.Name = $ReportSection.Section
                                    $Range = $Worksheet.Range("a1","$col$($temparray.GetLength(0))")
                                    $Range.Value2 = $temparray

                                    #Format the end result (headers, autofit, et cetera)
                                    [void]$Range.EntireColumn.AutoFit()
                                    [void]$Range.FormatConditions.Add($xlCellValue,$xlEqual,'TRUE')
                                    $Range.FormatConditions.Item(1).Interior.Color = $GoodColor
                                    $Range.FormatConditions.Item(1).Font.Color = $GoodText
                                    [void]$Range.FormatConditions.Add($xlCellValue,$xlEqual,'OK')
                                    $Range.FormatConditions.Item(2).Interior.Color = $GoodColor
                                    $Range.FormatConditions.Item(2).Font.Color = $GoodText
                                    [void]$Range.FormatConditions.Add($xlCellValue,$xlEqual,'FALSE')
                                    $Range.FormatConditions.Item(3).Interior.Color = $BadColor
                                    $Range.FormatConditions.Item(3).Font.Color = $BadText
                                    
                                    # Header
                                    $range = $Workbook.ActiveSheet.Range("a1","$($col)1")
                                    $range.Interior.ColorIndex = 19
                                    $range.Font.ColorIndex = 11
                                    $range.Font.Bold = $True
                                    $range.HorizontalAlignment = -4108
                                }
                            }
                        }
                        # Get rid of the blank default worksheets
                        $Workbook.Worksheets.Item("Sheet1").Delete()
                        $Workbook.Worksheets.Item("Sheet2").Delete()
                        $Workbook.Worksheets.Item("Sheet3").Delete()
                    }
                }
                #endregion Excel
            }

            foreach ($Asset in $AssetNames) {
                # First check if there is any data to report upon for each asset
                $ContainsData = $false
                $SectionCount = 0
                Foreach ($ReportSection in $SortedReports) {
                    if ($ReportContainer['Sections'][$ReportSection.Section]['AllData'].ContainsKey($Asset)) {
                        $ContainsData = $true
                    }
                }
                
                # If we have any data then we have a report to create
                if ($ContainsData) {
                    $AssetReport = ''
                    $AssetReport += $HTMLRendering['ServerBegin'][$HTMLMode] -replace '<0>',$Asset
                    $UsedSections = 0
                    $TotalSectionsPerRow = 0
                    
                    Foreach ($ReportSection in $SortedReports) {
                        if ($ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType]) {
                            #region Section Calculation
                            # Use this code to track where we are at in section usage
                            #  and create new section groups as needed
                            
                            # Current section type
                            $CurrContainer = $ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType]['ContainerType']
                            
                            # Grab first two digits found in the section container div
                            $SectionTracking = ([Regex]'\d{1}').Matches($HTMLRendering['SectionContainers'][$HTMLMode][$CurrContainer]['Head'])
                            if (($SectionTracking[1].Value -ne $TotalSectionsPerRow) -or `
                                ($SectionTracking[0].Value -eq $SectionTracking[1].Value) -or `
                                (($UsedSections + [int]$SectionTracking[0].Value) -gt $TotalSectionsPerRow) -and `
                                (!$ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType]['SectionOverride'])) {
                                    $NewGroup = $true
                            }
                            else {
                                $NewGroup = $false
                                $UsedSections += [int]$SectionTracking[0].Value
                            }
                            
                            if ($NewGroup) {
                                if ($UsedSections -ne 0) {
                                    $AssetReport += $HTMLRendering['SectionContainerGroup'][$HTMLMode]['Tail']
                                }
                                $AssetReport += $HTMLRendering['SectionContainerGroup'][$HTMLMode]['Head']
                                $UsedSections = [int]$SectionTracking[0].Value
                                $TotalSectionsPerRow = [int]$SectionTracking[1].Value
                            }
                            #endregion Section Calculation
                            $AssetReport += Create-ReportSection  -Rpt $ReportContainer `
                                                                  -Asset $Asset `
                                                                  -Section $ReportSection.Section `
                                                                  -TableTitle $ReportContainer['Sections'][$ReportSection.Section]['Title']
                        }
                    }
                    
                    $AssetReport += $HTMLRendering['SectionContainerGroup'][$HTMLMode]['Tail']
                    $AssetReport += $HTMLRendering['ServerEnd'][$HTMLMode]
                    $AssetReports += $AssetReport
                }
                # If we are creating per-asset reports then create one now, otherwise keep going
                if (($OutputMethod -eq 'IndividualReport') -and ($AssetReports -ne '')) {
                    $ReportOutputSplat.Report = ($HTMLRendering['Header'][$HTMLMode] -replace '<0>',$Asset) + 
                                                $AssetReports + 
                                                $HTMLRendering['Footer'][$HTMLMode]
                    $ReportOutputSplat.ReportName = $ReportNamePrefix + $Asset + '.html'
                    $ReportOutputSplat.ReportPath = $ReportLocation
            
                    $FinishedReportPath = New-ReportOutput @ReportOutputSplat
                    if ($FinishedReportPath -ne $false) {
                        $FinishedReportPaths += $FinishedReportPath
                    }
                    $AssetReports = ''
                }
            }
            
            # If one big report is getting sent/saved do so now
            if (($OutputMethod -eq 'OneBigReport') -and ($AssetReports -ne '')) {
                $FullReport = ($HTMLRendering['Header'][$HTMLMode] -replace '<0>',$Asset) + 
                                        $AssetReports + 
                                        $HTMLRendering['Footer'][$HTMLMode]
                $ReportOutputSplat.ReportName = $ReportName
                $ReportOutputSplat.ReportPath = $ReportLocation
                $ReportOutputSplat.Report = ($HTMLRendering['Header'][$HTMLMode] -replace '<0>','Multiple Systems') + 
                                                    $AssetReports + 
                                                    $HTMLRendering['Footer'][$HTMLMode]
                $FinishedReportPath = New-ReportOutput @ReportOutputSplat
                if ($FinishedReportPath -ne $false) {
                    $FinishedReportPaths += $FinishedReportPath
                }
            }
            
            if ($ZipReport) {
                $ZipReportName = "$($ReportOutputSplat.ReportName).zip"
                $FinishedReportPaths | Add-Zip $ZipReportName
                $FinishedReportPaths | Remove-Item
                $FinishedReportPaths = @($ZipReportName)
            }
            if ($SendMail) {
                $ReportDeliverySplat = @{
                    'EmailSender' = $EmailSender
                    'EmailRecipient' = $EmailRecipient
                    'EmailSubject' = $EmailSubject
                    'EmailRelay' = $EmailRelay
                    'SendMail' = $SendMail
                    'ForceAnonymous' = $ForceAnonymous
                }
                
                if ($ZipReport -or ($FinishedReportPaths.Count -gt 1)) {}
                New-ReportDelivery @ReportDeliverySplat
            }
        }
        else {
            Write-Verbose 'New-SelfContainedAssetReport: No data was able to be gathered!'
        }
    }
}