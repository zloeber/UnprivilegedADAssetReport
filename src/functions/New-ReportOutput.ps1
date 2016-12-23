Function New-ReportOutput
{
    [CmdletBinding()]
    param
    (
        [Parameter( HelpMessage="Report body, typically in HTML format",
                    ValueFromPipeline=$true,
                    Mandatory=$true )]
        [string]
        $Report,
        
        [Parameter( HelpMessage="Save the report as a PDF. If the PDF library is not available the default format, HTML, will be used instead.")]
        [switch]
        $SaveAsPDF,
        
        [Parameter( HelpMessage="Postpend timestamp to file name.")]
        [switch]
        $Postpendtimestamp,
        
        [Parameter( HelpMessage="Prepend timestamp to file name.")]
        [switch]
        $Prependtimestamp,
        
        [Parameter( HelpMessage="If output already exists do not overwrite.")]
        [switch]
        $NoOverwrite,
        
        [Parameter( HelpMessage="If saving the report, what do you want to call it?")]
        [string]
        $ReportName="Report.html",
        
        [Parameter( HelpMessage="Where are you saving the report (defaults to local temp directory)?")]
        [string]
        $ReportPath=$env:Temp
    )
    BEGIN
    {
        $timestamp = Get-Date -Format ddmmyyyy-HHMMss
        if ($Prependtimestamp)
        {
            $ReportName="$timestamp_$($ReportName.Split('.')[0]).$($ReportName.Split('.')[1])"
        }
        if ($Postpendtimestamp)
        {
            $ReportName="$($ReportName.Split('.')[0])_$timestamp.$($ReportName.Split('.')[1])"
        }
        $ReportFormat = 'HTML'
        if ($SaveAsPDF)
        {
            $PdfGenerator = "$((Get-Location).Path)\NReco.PdfGenerator.dll"
            if (Test-Path $PdfGenerator)
            {
                try {
                    $ReportFormat = 'PDF'
                    $PdfGenerator = "$((Get-Location).Path)\NReco.PdfGenerator.dll"
                    $Assembly = [Reflection.Assembly]::LoadFrom($PdfGenerator) #| Out-Null
                    $PdfCreator = New-Object NReco.PdfGenerator.HtmlToPdfConverter
                }
                catch {
                    $ReportFormat = 'HTML'
                }
            }
        }
    }
    PROCESS
    {}
    END
    {
        switch ($ReportFormat) {
            'PDF' {
                $ReportOutput = $PdfCreator.GeneratePdf([string]$Report)
                if ($ReportName -notmatch "\.pdf$") 
                {
                    if ($ReportName -match "\.html{0,1}$") 
                    {
                        $ReportName = [System.Text.RegularExpressions.Regex]::Replace($ReportName,"\.html{0,1}$", '.pdf');
                    }
                    else
                    {
                        $ReportName = "$($ReportName).pdf"
                    }
                }
                if ((Test-Path "$ReportPath\$ReportName") -and $NoOverwrite)
                {
                    $retval = $false
                }
                else
                {
                    Add-Content -Value $ReportOutput `
                                -Encoding byte `
                                -Path ("$ReportPath\$ReportName")
                    $retval = "$ReportPath\$ReportName"
                }
            }
            'HTML' {
                if ($ReportName -notmatch "\.html{0,1}$")
                {
                    if ($ReportName -match "\.pdf$") 
                    {
                        $ReportName = [System.Text.RegularExpressions.Regex]::Replace($ReportName,"\.pdf$", '.html');
                    }
                    else
                    {
                        $ReportName = "$($ReportName).html"
                    }
                }
                if ((Test-Path "$ReportPath\$ReportName") -and $NoOverwrite)
                {
                    $retval = $false
                }
                else
                {
                    $Report | Out-File "$ReportPath\$ReportName"
                    $retval = "$ReportPath\$ReportName"
                }
            }
        }
        return $retval
    }
}