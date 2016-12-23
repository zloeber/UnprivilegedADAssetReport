Function Get-ADPrivilegedGroups
{
    [CmdletBinding()]
    param
    (
        [Parameter(HelpMessage="Domain to gather privileged group information about. If not specified, all domains in the current forest will be enumerated.",
                   Mandatory=$false,
                   ValueFromPipeline=$true)]
        $Domain
    )
    BEGIN
    {
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
        Foreach ($Dom in $Domains)
        {
            # Domain SID
            $DomainDN = 'dc=' + $Dom.Replace('.', ',dc=')
            $DomGCobject = [adsi]"GC://$domainDN"
            $DomSid = New-Object System.Security.Principal.SecurityIdentifier($DomGCobject.objectSid[0], 0)
            $DomSid = $DomSid.toString()
            
            $StaticPrivGroupDesc = @{
                'S-1-5-32-544' = "Administrators"
                'S-1-5-32-548' = "Account Operators"
                'S-1-5-32-549' = "Server Operators"
                'S-1-5-32-550' = "Print Operators"
                'S-1-5-32-551' = "Backup Operators"
                "$DomSid-517" = "Cert Publishers"
                "$DomSid-518"  = "Schema Admins"
                "$DomSid-519"  = "Enterprise Admins"
               # "$DomSid-520"  = "Group Policy Creator Owners"
                "$DomSid-512"  = "Domain Admins"
            }
            $ADProp_Grp = @('Name',
                            'cn',
                            'distinguishedname')
            
            Foreach ($GrpSid in $StaticPrivGroupDesc.Keys)
            {
                $Grp = @(Search-AD -Filter "(objectSID=$GrpSid)" `
                                   -SearchRoot "LDAP://$DomainDN" `
                                   -Properties $ADProp_Grp)
                if ($Grp.Count -gt 0)
                {
                    $GrpProps = @{
                        'Domain' = $dom
                        'Group' = $StaticPrivGroupDesc[$GrpSid]
                        'GroupDN' = $Grp[0].distinguishedname
                        'GroupCN' = $Grp[0].cn
                        'GroupName' = $Grp[0].Name
                #        'Admincount' = $Grp[0].admincount
                        'Sid' = $GrpSid
                    }
                    New-Object PSObject -Property $GrpProps
                }
            }
        }
    }
}