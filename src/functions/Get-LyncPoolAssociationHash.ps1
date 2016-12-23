Function Get-LyncPoolAssociationHash 
{
    BEGIN
    {
        $Lync_Elements = @()
        $AD_PoolProperties = @('cn',
                               'distinguishedName',
                               'dnshostname',
                               'msrtcsip-pooldisplayname'
                              )
    }
    PROCESS
    {}
    END
    {
        $RootDSC = [adsi]"LDAP://RootDSE"
        $DomNamingContext = $RootDSC.RootDomainNamingContext
        $ConfigNamingContext = $RootDSC.configurationNamingContext
        $OCSADContainer = ''

        # Find Lync AD config partition 
        $LyncPathSearch = @(Search-AD -Filter '(objectclass=msRTCSIP-Service)' -SearchRoot "LDAP://$([string]$DomNamingContext)")
        if ($LyncPathSearch.count -ge 1)
        {
            $OCSADContainer = ($LyncPathSearch[0]).adspath
        }
        else
        {
            $LyncPathSearch = @(Search-AD -Filter '(objectclass=msRTCSIP-Service)' -SearchRoot "LDAP://$ConfigNamingContext")
            if ($LyncPathSearch.count -ge 1)
            {
                $OCSADContainer = ($LyncPathSearch[0]).adspath
            }
        }
        if ($OCSADContainer -ne '')
        {
            $LyncPoolLookupTable = @{}
            # All Lync pools
            $Lync_Pools = @(Search-AD -Filter '(&(objectClass=msRTCSIP-Pool))' `
                                      -Properties $AD_PoolProperties `
                                      -SearchRoot $OCSADContainer)
            $LyncPoolCount = $Lync_Pools.Count
            $Lync_Pools | %{
                $LyncElementProps = @{
                    CN = $_.cn
                    distinguishedName = $_.distinguishedName
                    ServiceName = "CN=Lc Services,CN=Microsoft,$($_.distinguishedName)"
                    PoolName = $_.'msrtcsip-pooldisplayname'
                    PoolFQDN = $_.dnshostname
                }
                $Lync_Elements += New-Object PSObject -Property $LyncElementProps
            }
            $Lync_Elements
        }
    }
}