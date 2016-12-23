Function Get-ADDomainPrivAccounts
{
    [CmdletBinding()]
    param
    (
        [Parameter(HelpMessage="Domain to gather privileged accounts. If not specified, all domains in the current forest will be enumerated.",
                   ValueFromPipeline=$true)]
        [string[]]$Domain,
        [Parameter(HelpMessage='User attributes to include in results.')]
        $UserAttribs = @( 'cn',
                          'displayName',
                          'givenName',
                          'sn',
                          'name',
                          'sAMAccountName',
                          'whenChanged',
                          'whenCreated',
                          'pwdLastSet',
                          'badPasswordTime',
                          'badPwdCount',
                          'lastLogon',
                          'logonCount',
                          'useraccountcontrol',
                          'lastlogontimestamp'
                        )
    )
    BEGIN
    {
        $RootDSC = [adsi]"LDAP://RootDSE"
        $DomNamingContext = $RootDSC.RootDomainNamingContext
        $ConfigNamingContext = $RootDSC.configurationNamingContext
        $Domains = @()
    }
    PROCESS
    {
        if ($Domain -ne $null)
        {
            $Domains += $Domain
        }
    }
    END
    {
        if ($Domains.Count -eq 0)
        {
            $Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            $Domains = @($Forest.Domains | %{[string]$_.Name})
        }
        $DomPrivGroups = @()
        ForEach ($Dom in $Domains) 
        { 
            $CurDomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $Dom)
            $CurDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($CurDomainContext)
            $CurDomainDetails = [ADSI]"LDAP://$($CurDomain)"
            $DomainDN = 'dc=' + $Dom.Replace('.', ',dc=')
            $NetBIOSName = Get-NETBiosName $DomainDN $ConfigNamingContext
            
            $DomPrivGroups = @(Get-ADPrivilegedGroups -Domain $Dom)
            Foreach ($PrivGroup in $DomPrivGroups)
            {
                $PrivGroupDN = $PrivGroup.GroupDN
                Write-Verbose $PrivGroupDN
                # Only works on 2003 SP2 and above
                $Filter = "(samAccountType=805306368)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(memberOf:1.2.840.113556.1.4.1941:=$PrivGroupDN)"
                $PrivUsers = @(Search-AD -Filter $Filter `
                                         -SearchRoot "LDAP://$DomainDN" `
                                         -Properties $UserAttribs)
                Write-Verbose -Message ('Privileged Users: Group {0}' -f $PrivGroup.GroupDN)
                $PrivUsers = $PrivUsers | 
                             Normalize-ADUsers -Attribs $UserAttribs |
                             Append-ADUserAccountControl
                Foreach ($PrivUser in $PrivUsers)
                {
                    if ($PrivUser -ne $null)
                    {
                        $PrivMemberProp = @{
                            Domain = $Dom
                            DomainNetBIOS = $NetBIOSName
                            PrivGroup = $PrivGroup.Group
                        }
                        $PrivUser.psobject.properties | 
                        Where {$_.Name -ne $null} | ForEach {
                            $PrivMemberProp[$_.Name] = $_.Value 
                        }
                        New-Object psobject -Property $PrivMemberProp
                    }
                }
            }
        }
    }
}