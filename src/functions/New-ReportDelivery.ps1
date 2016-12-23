Function New-ReportDelivery
{
    [CmdletBinding()]
    param
    (
        [Parameter( HelpMessage="Report body, typically in HTML format", ValueFromPipeline=$true )]
        [string[]]
        $Report,
        
        [Parameter( ParameterSetName="EmailReport", HelpMessage="Send email of resulting report?")]
        [Parameter( ParameterSetName = "EmailAndSaveReport")]                    
        [switch]
        $SendMail,
        
        [Parameter( ParameterSetName="EmailReport", HelpMessage="Email server to relay report through")]
        [Parameter( ParameterSetName = "EmailAndSaveReport")]
        [string]
        $EmailRelay = ".",
        
        [Parameter( ParameterSetName="EmailReport", HelpMessage="Email sender")]
        [Parameter( ParameterSetName = "EmailAndSaveReport")]
        [string]
        $EmailSender='systemreport@localhost',
        
        [Parameter( ParameterSetName="EmailReport", Mandatory=$true, HelpMessage="Email recipient")]
        [Parameter( ParameterSetName = "EmailAndSaveReport")]
        [string]
        $EmailRecipient,
        
        [Parameter( ParameterSetName="EmailReport", HelpMessage="Email subject")]
        [Parameter( ParameterSetName = "EmailAndSaveReport")]
        [string]
        $EmailSubject='System Report',
        
        [Parameter( ParameterSetName="EmailReport", HelpMessage="Email report(s) as attachement")]
        [Parameter( ParameterSetName = "EmailAndSaveReport")]
        [Parameter( ParameterSetName = "EmailReportAsAttachment")]
        [switch]
        $EmailAsAttachment,
        
        [Parameter( ParameterSetName="EmailReport", HelpMessage="Force email to be sent anonymously?")]
        [Parameter( ParameterSetName = "EmailAndSaveReport")]
        [switch]
        $ForceAnonymous,

        [Parameter( ParameterSetName="SaveReport", HelpMessage="Save the report?")]
        [Parameter( ParameterSetName = "EmailAndSaveReport")]
        [switch]
        $SaveReport,
        
        [Parameter( ParameterSetName="SaveReport", HelpMessage="Zip the report(s).")]
        [Parameter( ParameterSetName = "EmailAndSaveReport")]
        [Parameter( ParameterSetName = "EmailReportAsAttachment")]
        [switch]
        $ZipReport
    )
    BEGIN
    {
        $Reports = @()      # Save a list of report paths in case we will be emailing as attachments
        if ($SaveReport)
        {
            $ReportFormat = 'HTML'
        }
        if ($SaveAsPDF)
        {
            $PdfGenerator = "$((Get-Location).Path)\NReco.PdfGenerator.dll"
            if (Test-Path $PdfGenerator)
            {
                $ReportFormat = 'PDF'
                $PdfGenerator = "$((Get-Location).Path)\NReco.PdfGenerator.dll"
                $Assembly = [Reflection.Assembly]::LoadFrom($PdfGenerator) #| Out-Null
                $PdfCreator = New-Object NReco.PdfGenerator.HtmlToPdfConverter
            }
        }
    }
    PROCESS
    {
        switch ($ReportFormat) {
            'PDF' {
                $ReportOutput = $PdfCreator.GeneratePdf([string]$Report)
                $ReportName = $ReportName -replace '.html','.pdf'
                Add-Content -Value $ReportOutput `
                            -Encoding byte `
                            -Path ($ReportName)
            }
            'HTML' {
                $Report | Out-File $ReportName
            }
        }
        $Reports += $ReportName
    }
    END
    {
        if ($Sendmail)
        {
            $SendMailSplat = @{
                'From' = $EmailSender
                'To' = $EmailRecipient
                'Subject' = $EmailSubject
                'Priority' = 'Normal'
                'smtpServer' = $EmailRelay
                'BodyAsHTML' = $true
            }
            if ($ForceAnonymous)
            {
                $Pass = ConvertTo-SecureString –String 'anonymous' –AsPlainText -Force
                $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "NT AUTHORITY\ANONYMOUS LOGON", $pass
                $SendMailSplat.Credential = $creds

            }
            if ($EmailAsAttachment)
            {
                if ($ZipReport)
                {
                    $ZipName = $ReportName -replace '.html','.zip'
                    $Reports | New-ZipFile -ZipFilePath $ZipName -Append
                }
                else
                {
                    $SendMailSplat.Attachments = $Reports
                }
            }
            else
            {
                $SendMailSplat.Body = $Report
            }
            send-mailmessage @SendMailSplat
        }
    }
}
