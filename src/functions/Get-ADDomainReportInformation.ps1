Function Get-ADDomainReportInformation
{
    [CmdletBinding()]
    param
    (
        [Parameter( HelpMessage="The custom report hash variable structure you plan to report upon")]
        $ReportContainer,
        [Parameter( HelpMessage="A sorted hash of enabled report elements.")]
        $SortedRpts
    )
    BEGIN
    {
        try
        {
            $verbose_timer = Get-Date
            $Filter_Users = '(samAccountType=805306368)'
            $Filter_User_Locked = '(samAccountType=805306368)(lockoutTime:1.2.840.113556.1.4.804:=4294967295)'
            $Filter_User_PasswordChangeReq = '(samAccountType=805306368)(pwdLastSet=0)(!useraccountcontrol:1.2.840.113556.1.4.803:=2)'
            $Filter_User_Enabled = '(samAccountType=805306368)(!(userAccountControl:1.2.840.113556.1.4.803:=2))'
            $Filter_User_Disabled = '(samAccountType=805306368)(useraccountcontrol:1.2.840.113556.1.4.803:=2)'
            $Filter_User_NoPasswordReq = '(samAccountType=805306368)(UserAccountControl:1.2.840.113556.1.4.803:=32)'
            $Filter_User_PasswordNeverExpires = '(samAccountType=805306368)(UserAccountControl:1.2.840.113556.1.4.803:=65536)'
            $Filter_User_DialinEnabled = '(samAccountType=805306368)(msNPAllowDialin=TRUE)'
            $Filter_User_UnconstrainedDelegation = '(samAccountType=805306368)(userAccountControl:1.2.840.113556.1.4.803:=524288)'
            $Filter_User_NotTrustedForDelegation = '(samAccountType=805306368)(userAccountControl:1.2.840.113556.1.4.803:=524288)'
            $Filter_User_NoPreauth = '(samAccountType=805306368)(userAccountControl:1.2.840.113556.1.4.803:=4194304)'
            $Filter_User_ControlAccessWithNPS = '(samAccountType=805306368)(!(msNPAllowDialin=*))'

            $RootDSC = [adsi]"LDAP://RootDSE"
            $DomNamingContext = $RootDSC.RootDomainNamingContext
            $ConfigNamingContext = $RootDSC.configurationNamingContext
            $Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            $Domains = @($Forest.Domains | %{[string]$_.Name})
            
            $ADConnected = $true
        }
        catch
        {
            $ADConnected = $false
        }
    }
    PROCESS
    {}
    END
    {
        if ($ADConnected)
            {
            Foreach ($Dom in $Domains)
            {
                $CurDomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $Dom)
                try
                {
                    $CurDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($CurDomainContext)
                    Write-Verbose -Message ('Get-ADDomainReportInformation: Start - {0}' -f $verbose_timer)
                    Write-Verbose -Message ('Get-ADDomainReportInformation: Domain - {0}' -f $Dom)
                    $UserStats = $null
                    $GroupStats = $null
                    $PrivGroups = $null
                    $PrivGroupMembers = $null
                    $TotalPrivGroupCount = 0
                    
                    $DomainDN = 'dc=' + $Dom.Replace('.', ',dc=')
                    $Splat_SearchAD = @{
                        'SearchRoot' = "LDAP://$DomainDN"
                        'Properties' = $UserAttribs
                    }
                    if ($ExportAllUsers)
                    {
                        Write-Verbose -Message ('Get-ADDomainReportInformation: Export all users in domain - {0}' -f $Dom)
                        Search-AD -Properties $UserAttribs `
                                  -Filter '(samAccountType=805306368)' `
                                  -SearchRoot "LDAP://$DomainDN" |
                            Normalize-ADUsers -Attribs $UserAttribs | 
                                Append-ADUserAccountControl |
                                    Export-Csv -NoTypeInformation "allusers_$Dom.csv"
                        Write-Verbose -Message ('Get-ADDomainReportInformation: Timer - {0}' -f $((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    }
                    Write-Verbose -Message ('Get-ADDomainReportInformation: Domain User Stats - {0}' -f $Dom)
                    $UserStats = New-Object psobject -Property @{
                        'Total' = @(Search-AD @Splat_SearchAD -Filter $Filter_Users).Count
                        'Enabled' = @(Search-AD @Splat_SearchAD -Filter $Filter_User_Enabled).Count
                        'Disabled' = @(Search-AD @Splat_SearchAD -Filter $Filter_User_Disabled).Count
                        'Locked' = @(Search-AD @Splat_SearchAD -Filter $Filter_User_Locked).Count
                        'PwdDoesNotExpire' = @(Search-AD @Splat_SearchAD -Filter $Filter_User_PasswordNeverExpires).Count
                        'PwdNotRequired' = @(Search-AD @Splat_SearchAD -Filter $Filter_User_NoPasswordReq).Count
                        'PwdMustChange' = @(Search-AD @Splat_SearchAD -Filter $Filter_User_PasswordChangeReq).Count
                        'DialInEnabled' = @(Search-AD @Splat_SearchAD -Filter $Filter_User_DialinEnabled).Count
                        'UnconstrainedDelegation' = @(Search-AD @Splat_SearchAD -Filter $Filter_User_UnconstrainedDelegation).Count
                        'NotTrustedForDelegation' = @(Search-AD @Splat_SearchAD -Filter $Filter_User_NotTrustedForDelegation).Count
                        'NoPreAuthRequired' = @(Search-AD @Splat_SearchAD -Filter $Filter_User_NoPreauth).Count
                        'ControlAccessWithNPS' = @(Search-AD @Splat_SearchAD -Filter $Filter_User_ControlAccessWithNPS).Count
                    }
                    Write-Verbose -Message ('Get-ADDomainReportInformation: Timer - {0}' -f $((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    $AllGroups = @(
                        Search-AD -Properties groupType `
                                  -Filter '(objectClass=group)' `
                                  -SearchRoot "LDAP://$DomainDN"
                    )
                    Write-Verbose -Message ('Get-ADDomainReportInformation: Domain Group Stats - {0}' -f $Dom)
                    $GroupStats = New-Object psobject -Property @{
                        'Total' = $AllGroups.Count
                        'Builtin' = @($AllGroups | Where {$_.groupType -eq '-2147483643'}).Count
                        'UniversalSecurity' = @($AllGroups | Where {$_.groupType -eq '-2147483640'}).Count
                        'UniversalDist' = @($AllGroups | Where {$_.groupType -eq '8'}).Count
                        'GlobalSecurity' = @($AllGroups | Where {$_.groupType -eq '-2147483646'}).Count
                        'GlobalDist' = @($AllGroups | Where {$_.groupType -eq '2'}).Count
                        'DomainLocalSecurity' = @($AllGroups | Where {$_.groupType -eq '-2147483644'}).Count
                        'DomainLocalDist' = @($AllGroups | Where {$_.groupType -eq '4'}).Count
                    }
                    $PrivGroups = @(Get-ADPrivilegedGroups -Domain $Dom)
                    $PrivUsers = @(Get-ADDomainPrivAccounts -Domain $Dom)
                    Write-Verbose -Message ('Get-ADDomainReportInformation: Timer - {0}' -f $((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    if ($ExportPrivilegedUsers)
                    {
                        Write-Verbose -Message ('Get-ADDomainReportInformation: Exporting privileged users - {0}' -f $Dom)
                        $PrivUsers | Export-Csv -NoTypeInformation "privusers_$Dom.csv"
                        Write-Verbose -Message ('Get-ADDomainReportInformation: Timer - {0}' -f $((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    }
                    $PrivGroupStats = @()
                    
                    ForEach ($PrivGroup in $PrivGroups)
                    {
                        Foreach ($PrivGrp in $AD_PrivilegedGroups)
                        {
                            if ($PrivGrp -eq $PrivGroup.Group)
                            {
                                $PrivGroupCount = @($PrivUsers | Where {$_.PrivGroup -eq $PrivGrp}).Count
                                $TotalPrivGroupCount = $TotalPrivGroupCount + $PrivGroupCount
                                $PrivGroupStatProp = @{
                                    AdminGroup =  $PrivGrp
                                    DisplayName = $PrivGroup.GroupName
                                    MemberCount = $PrivGroupCount
                                }
                                $PrivGroupStats += New-Object psobject -Property $PrivGroupStatProp
                            }
                        }
                    }
                    #region Populate Data
                    $SortedRpts | %{ 
                        switch ($_.Section) {
                            'UserAccountStats1' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($UserStats)
                            }
                            'UserAccountStats2' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($UserStats)
                            }
                            'GroupStats' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($GroupStats)
                            }
                            'PrivGroupStats' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($PrivGroupStats)
                            }
                            'PrivGroup_EnterpriseAdmins' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($PrivUsers | 
                                      Where {$_.PrivGroup -eq 'Enterprise Admins'} |
                                      Sort-Object -Property PasswordAge -Descending)
                            }
                            'PrivGroup_SchemaAdmins' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($PrivUsers | 
                                      Where {$_.PrivGroup -eq 'Schema Admins'} |
                                      Sort-Object -Property PasswordAge -Descending)
                            }
                            'PrivGroup_DomainAdmins' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($PrivUsers | 
                                      Where {$_.PrivGroup -eq 'Domain Admins'} |
                                      Sort-Object -Property PasswordAge -Descending)
                            }
                            'PrivGroup_Administrators' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($PrivUsers | 
                                      Where {$_.PrivGroup -eq 'Administrators'} |
                                      Sort-Object -Property PasswordAge -Descending)
                            }
                            'PrivGroup_AccountOperators' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($PrivUsers | 
                                      Where {$_.PrivGroup -eq 'Account Operators'} |
                                      Sort-Object -Property PasswordAge -Descending)
                            }
                            'PrivGroup_ServerOperators' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($PrivUsers | 
                                      Where {$_.PrivGroup -eq 'Server Operators'} |
                                      Sort-Object -Property PasswordAge -Descending)
                            }
                            'PrivGroup_BackupOperators' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($PrivUsers | 
                                      Where {$_.PrivGroup -eq 'Backup Operators'} |
                                      Sort-Object -Property PasswordAge -Descending)
                            }
                            'PrivGroup_PrintOperators' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($PrivUsers | 
                                      Where {$_.PrivGroup -eq 'Print Operators'} |
                                      Sort-Object -Property PasswordAge -Descending)
                            }
                            'PrivGroup_CertPublishers' {
                                $ReportContainer['Sections'][$_]['AllData'][$Dom] = 
                                    @($PrivUsers | 
                                      Where {$_.PrivGroup -eq 'Cert Publishers'} |
                                      Sort-Object -Property PasswordAge -Descending)
                            }
                        }
                    }
                    #endregion Populate Data
                }
                catch
                {
                    Write-Warning ('Get-ADForestReportInformation: Issue with {0} Domain - {1}' -f $Dom,$_.Exception.Message)
                }
            }
            $ReportContainer['Configuration']['Assets'] = $Domains
            Return $Domains
        }
    }
}