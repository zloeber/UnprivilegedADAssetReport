$ADForestReportPreProcessing =
@'
    Get-ADForestReportInformation @VerboseDebug `
                               -ReportContainer $ReportContainer `
                               -SortedRpts $SortedReports
'@

$ADDomainReportPreProcessing =
@'
    Get-ADDomainReportInformation @VerboseDebug `
                               -ReportContainer $ReportContainer `
                               -SortedRpts $SortedReports
'@

$LyncElements_Postprocessing =
@'
    $temp = Format-HTMLTable $Table -Column 'Type' -ColumnValue 'Internal' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'Type' -ColumnValue 'Backend' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'Type' -ColumnValue 'Pool' -Attr 'class' -AttrValue 'warn'
            Format-HTMLTable $temp -Column 'Type' -ColumnValue 'Edge' -Attr 'class' -AttrValue 'alert'
'@

$ForestDomainDNSZones_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -match [string]$args[1]}
    $temp = Format-HTMLTable $Table -Scriptblock $scriptblock -Column 'Name' -ColumnValue 'CNF:' -Attr 'class' -AttrValue 'warn'
            Format-HTMLTable $temp  -Scriptblock $scriptblock -Column 'Name' -ColumnValue 'InProgress' -Attr 'class' -AttrValue 'warn'
'@
$ForestSiteConnections_Postprocessing =
@'
    $temp = Format-HTMLTable $Table -Column 'Enabled' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
            Format-HTMLTable $temp -Column 'Enabled' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
'@

$ForestDomainDCs_Postprocessing =
@'
    $temp = Format-HTMLTable $Table -Column 'GC' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'GC' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp -Column 'Infra' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'Infra' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp -Column 'Naming' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'Naming' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp -Column 'Schema' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'Schema' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp -Column 'RID' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'RID' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp -Column 'PDC' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
            Format-HTMLTable $temp -Column 'PDC' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
'@

$ADPrivUser_Postprocessing =
@'
    [scriptblock]$scriptblock = {[int]$args[0] -ge [int]$args[1]}
    [scriptblock]$scriptblockhealthy = {[int]$args[0] -lt [int]$args[1]}
    $temp = Format-HTMLTable $Table -Column 'No Pwd Expiry' -ColumnValue 'True' -Attr 'class' -AttrValue 'warn'
    $temp = Format-HTMLTable $temp -Column 'No Pwd Expiry' -ColumnValue 'False' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'Pwd Reversable' -ColumnValue 'True' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp -Column 'Pwd Reversable' -ColumnValue 'False' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'Pwd Not Req.' -ColumnValue 'True' -Attr 'class' -AttrValue 'warn'
    $temp = Format-HTMLTable $temp -Column 'Pwd Not Req.' -ColumnValue 'False' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Scriptblock $scriptblockhealthy -Column 'Pwd Age (Days)' -ColumnValue $AD_PwdAgeHealthy -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Scriptblock $scriptblock        -Column 'Pwd Age (Days)' -ColumnValue $AD_PwdAgeWarn -Attr 'class' -AttrValue 'warn'
            Format-HTMLTable $temp -Scriptblock $scriptblock        -Column 'Pwd Age (Days)' -ColumnValue $AD_PwdAgeAlert -Attr 'class' -AttrValue 'alert'
'@