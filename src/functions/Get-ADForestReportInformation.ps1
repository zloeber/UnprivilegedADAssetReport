Function Get-ADForestReportInformation {
    [CmdletBinding()]
    param (
        [Parameter( HelpMessage="The custom report hash variable structure you plan to report upon")]
        $ReportContainer,
        [Parameter( HelpMessage="A sorted hash of enabled report elements.")]
        $SortedRpts
    )
    Begin {
        $verbose_timer = $verbose_starttime = Get-Date
        $ldapregex = [regex]'(?<LDAPType>^.+)\=(?<LDAPName>.+$)'
        try {
            $ADConnected = $true
            $schema = [DirectoryServices.ActiveDirectory.ActiveDirectorySchema]::GetCurrentSchema()
            $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            $GCs = $forest.FindAllGlobalCatalogs()
            $GCNames = @($GCs | Select-Object Name)
            $ForestDCs = @($forest.Domains | Foreach-Object {$_.DomainControllers} | Select-Object Name)
            $ForestGCs = @((($GCs | Sort-Object -Property Name) | Select-Object Name))
            $schemapartition = $schema.Name
            $RootDSC = [adsi]"LDAP://RootDSE"
            $DomNamingContext = $RootDSC.RootDomainNamingContext
            $ConfigNamingContext = $RootDSC.configurationNamingContext
            
            # Start assuming Lync isn't configured in the environment
            $Lync_ConfigPartition = 'None'

            # Based on our connection thus far create a bunch of LDAP paths for use searching later
            $Path_LDAPPolicies = "LDAP://CN=Default Query Policy,CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services,$($ConfigNamingContext)"
            $Path_RecycleBinFeature = "LDAP://CN=Recycle Bin Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,$($ConfigNamingContext)"
            $Path_TombstoneLifetime = "LDAP://CN=Directory Service,CN=Windows NT,CN=Services,$($ConfigNamingContext)"
            $Path_ExchangeOrg = "LDAP://CN=Microsoft Exchange,CN=Services,$($ConfigNamingContext)"
            $Path_ExchangeVer = "LDAP://CN=ms-Exch-Schema-Version-Pt,$($SchemaPartition)"
            $Path_LyncVer = "LDAP://CN=ms-RTC-SIP-SchemaVersion,$($SchemaPartition)"
            $Path_ADSubnets = "LDAP://CN=Subnets,CN=Sites,$($ConfigNamingContext)"
            $Path_ADSiteLinks = "LDAP://CN=Sites,$($ConfigNamingContext)"
            
            # Initialize a bunch of stuff to be filled out later
            $ExchangeFederations = @()
            $ExchangeServers = @()
            $Lync_Elements = @()
            $Sites = @()
            $SiteSubnets = @()
            $AllSiteConnections = @()
            $SiteLinks = @()
            $DomainControllers = @()
            $Domains = @()
            $DomainDFS = @()
            $DomainDFSR = @()
            $DomainTrusts = @()
            $DomainDNSZones = @()
            $DomainGPOs = @()
            $NPSServers = @()
            $DomainPrinters = @()
            $DomainPrivGroups = @()
        }
        catch {
            $ADConnected = $false
        }
    }
    Process {}
    End {
        if ($ADConnected) {
            Write-Verbose -Message ('Get-ADForestReportInformation {0}: Forest Info - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
            #region Forest Settings
            # Recycle Bin Feature check
            $RecycleBinEnabled = $false
            if ([ADSI]::Exists($Path_RecycleBinFeature)) {
                $RecycleBinAttribs = Search-AD -Properties * -SearchRoot $Path_RecycleBinFeature
                if ($RecycleBinAttribs.PSObject.Properties.Match('msDS-EnabledFeatureBL').Count) 
                {
                    $RecycleBinEnabled = $True
                }
            }

            if ([ADSI]::Exists($Path_TombstoneLifetime)) {
                [ADSI]$TombstoneConfig = $Path_TombstoneLifetime
                $TombstoneLife = $TombstoneConfig.TombstoneLifetime
                $DeletedObjectLife = $TombstoneConfig."msDS-DeletedObjectLifetime"
                if ($TombstoneLife -ne $null) {
                    $TotalObjectBackupLife = $TombstoneLife
                }
                if ($deletedObjectLife) {
                    if ((-not $TombstoneLife) -or ($DeletedObjectLife -lt $TombstoneLife)) {
                        $TotalObjectBackupLife = $deletedObjLifetime
                    }
                }
            }
            else {
                $TombstoneLife = 'NA'
                $DeletedObjectLife = 'NA'
                $TotalObjectBackupLife = 'NA'
            }
            if ([ADSI]::Exists($Path_LDAPPolicies)) {
                [ADSI]$LDAPPoliciesConfig = $Path_LDAPPolicies
                $LDAPAdminLimits = $LDAPPoliciesConfig.LDAPAdminLimits
            }
            else {
                $LDAPAdminLimits = $null
            }
            #endregion Forest Settings
            
            #region DHCP Servers
            $DHCPServers = @(Search-AD -Filter '(objectclass=dHCPClass)' -Properties Name,WhenCreated -SearchRoot "LDAP://$([string]$ConfigNamingContext)" | Where-Object {$_.Name -ne 'DhcpRoot'})
            #endregion DHCP Servers
            
            #region Exchange

            $ExchangeServerCount = 0
            if ([ADSI]::Exists($Path_ExchangeVer)) {
                Write-Verbose -Message ('Get-ADForestReportInformation {0}: Exchange - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                [ADSI]$SchemaPathExchange = $Path_ExchangeVer
                $ExchangeSchema = ($SchemaPathExchange | Select-Object rangeUpper).rangeUpper
                $ExchangeVersion = $SchemaHashExchange[$ExchangeSchema]
                $Props_ExchOrgs = @('distinguishedName',
                                    'Name')
                $Props_ExchServers = @('adspath',
                                       'Name',
                                       'msexchserversite',
                                       'msexchcurrentserverroles',
                                       'adminDisplayName',
                                       'whencreated',
                                       'serialnumber',
                                       'msexchproductid')
                $Props_ExchFeds = @('Name',
                                    'msExchFedIsEnabled',
                                    'msExchFedDomainNames',
                                    'msExchFedEnabledActions',
                                    'msExchFedTargetApplicationURI',
                                    'msExchFedTargetAutodiscoverEPR',
                                    'msExchVersion')

                if ([ADSI]::Exists($Path_ExchangeOrg)) {
                    $ExchOrgs = @(Search-AD -Filter '(&(objectClass=msExchOrganizationContainer))' `
                                            -Properties $Props_ExchOrgs `
                                            -SearchRoot $Path_ExchangeOrg)
                    foreach ($ExchOrg in $ExchOrgs) {
                        $ExchServers = @(Search-AD -Filter '(objectCategory=msExchExchangeServer)' `
                                                   -Properties $Props_ExchServers `
                                                   -SearchRoot "LDAP://$([string]$ExchOrg.distinguishedname)")
                        $ExchangeServerCount += $ExchServers.Count
                        foreach ($ExchServer in $ExchServers)
                        {
                            $AdminGroup = Get-ADPathName $ExchServer.adspath -GetElement 2 -ValuesOnly
                            $ExchSite =  Get-ADPathName $ExchServer.msexchserversite -GetElement 0 -ValuesOnly
                            $ExchRole = $ExchServer.msexchcurrentserverroles
                            # only have two roles in Exchange 2013 so we process a bit differently
                            if ($ExchServer.serialNumber -like "Version 15*")
                            {
                                switch ($ExchRole) {
                                    '54' {
                                        $ExchRole = 'MAILBOX'
                                    }
                                    '16385' {
                                        $ExchRole = 'CAS'
                                    }
                                    '16439' {
                                        $ExchRole = 'MAILBOX, CAS'
                                    }
                                }
                            }
                            else
                            {
                                if($ExchRole -ne 0)
                                {
                                    $ExchRole = [Enum]::Parse('MSExchCurrentServerRolesFlags', $ExchRole)
                                }
                            }
                            $exchserverprops = @{
                                Organization = $ExchOrg.Name
                                AdminGroup   = $AdminGroup
                                Name         = $ExchServer.adminDisplayName
                                Role         = $ExchRole
                                Site         = $ExchSite
                                Created      = $ExchServer.whencreated
                                Serial       = $ExchServer.serialnumber
                                ProductID    = $ExchServer.msexchproductid
                            }
                            $ExchangeServers += New-Object PSObject -Property $exchserverprops
                        }
                        $ExchangeFeds = @(Search-AD -Filter '(objectCategory=msExchFedSharingRelationship)' `
                                                   -Properties $Props_ExchFeds -DontJoinAttributeValues `
                                                   -SearchRoot "LDAP://CN=Federation,$([string]$ExchOrg.distinguishedname)")
                        Foreach ($ExchFed in $ExchangeFeds)
                        {
                            $ExchangeFedProps = @{
                                Organization = $ExchOrg.Name
                                Name = $ExchFed.Name
                                Enabled = $ExchFed.msExchFedIsEnabled
                                Domains = @($ExchFed.msExchFedDomainNames)
                                AllowedActions = @($ExchFed.msExchFedEnabledActions)
                                TargetAppURI = $ExchFed.msExchFedTargetApplicationURI
                                TargetAutodiscoverEPR = $ExchFed.msExchFedTargetAutodiscoverEPR
                                ExchangeVersion = $ExchFed.msExchVersion
                            }
                            $ExchangeFederations += New-Object psobject -Property $ExchangeFedProps
                        }
                    }
                }
            }
            else
            {
                $ExchangeVersion = 'Exchange Not Installed'
            }
            #endregion Exchange

            #region OCS/Lync
            $Lync_InternalServers = @()
            $Lync_EdgeServers = @()
            $Lync_Pools = @()
            $Lync_OtherServers = @()
            $LyncServerCount = 0
            $LyncPoolCount = 0
            if([ADSI]::Exists($Path_LyncVer))
            {
                Write-Verbose -Message ('Get-ADForestReportInformation {0}: Lync/OCS - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                # Get Lync version in forest
                [ADSI]$SchemaPathLync = $Path_LyncVer
                $LyncSchema = ($SchemaPathLync | Select-Object rangeUpper).rangeUpper
                $LyncVersion = $SchemaHashLync[$LyncSchema]
                
                # Find Lync AD config partition location
                $LyncPathSearch = @(Search-AD -Filter '(objectclass=msRTCSIP-Service)' -SearchRoot "LDAP://$([string]$DomNamingContext)")
                if ($LyncPathSearch.count -ge 1)
                {
                    $OCSADContainer = ($LyncPathSearch[0]).adspath
                    $Lync_ConfigPartition = 'System'
                }
                else
                {
                    $LyncPathSearch = @(Search-AD -Filter '(objectclass=msRTCSIP-Service)' -SearchRoot "LDAP://$ConfigNamingContext")
                    if ($LyncPathSearch.count -ge 1)
                    {
                        $OCSADContainer = ($LyncPathSearch[0]).adspath
                        $Lync_ConfigPartition = 'Configuration'
                    }
                }
                
                # All internal Lync servers
                Search-AD -Filter '(&(objectClass=msRTCSIP-TrustedServer))' `
                          -Properties 'msrtcsip-trustedserverfqdn',Name `
                          -SearchRoot $OCSADContainer | 
                Sort-Object msrtcsip-trustedserverfqdn | %{ 
                    $LyncElementProps = @{
                        LyncElement = 'Server'
                        LyncElementType = 'Internal'
                        LyncElementName = $_.Name
                        LyncElementFQDN = $_.'msrtcsip-trustedserverfqdn'
                    }
                    $Lync_Elements += New-Object PSObject -Property $LyncElementProps
                }
                # All edge Lync servers
                Search-AD -Filter '(&(objectClass=msRTCSIP-EdgeProxy))' `
                          -Properties cn,Name,'msrtcsip-edgeproxyfqdn' `
                          -SearchRoot $OCSADContainer | 
                Sort-Object msrtcsip-edgeproxyfqdn | %{
                    $LyncElementProps = @{
                        LyncCN = $_.cn
                        LyncElement = 'Server'
                        LyncElementType = 'Edge'
                        LyncElementName = $_.Name
                        LyncElementFQDN = $_.'msrtcsip-edgeproxyfqdn'
                    }
                    $Lync_Elements += New-Object PSObject -Property $LyncElementProps
                }
                # All Lync global topology servers
                Search-AD -Filter '(&(objectClass=msRTCSIP-GlobalTopologySetting))' `
                          -Properties cn,Name,'msrtcsip-backendserver' `
                          -SearchRoot $OCSADContainer | 
                Sort-Object msrtcsip-backendserver | %{
                    $LyncElementProps = @{
                        LyncCN = $_.cn
                        LyncElement = 'Server'
                        LyncElementType = 'Backend'
                        LyncElementName = $_.Name
                        LyncElementFQDN = $_.'msrtcsip-backendserver'
                    }
                    $Lync_Elements += New-Object PSObject -Property $LyncElementProps
                }
                
                $LyncServerCount = $Lync_Elements.Count
                
                # All Lync pools
                $Lync_Pools = @(Search-AD -Filter '(&(objectClass=msRTCSIP-Pool))' `
                                          -Properties cn,dnshostname,'msrtcsip-pooldisplayname' `
                                          -SearchRoot $OCSADContainer | 
                                    Sort-Object dnshostname)
                $LyncPoolCount = $Lync_Pools.Count
                $Lync_Pools | %{
                    $LyncElementProps = @{
                        LyncCN = $_.cn
                        LyncElement = 'Pool'
                        LyncElementType = 'Pool'
                        LyncElementName = $_.'msrtcsip-pooldisplayname'
                        LyncElementFQDN = $_.dnshostname
                    }
                    $Lync_Elements += New-Object PSObject -Property $LyncElementProps
                }
            }
            else
            {
                $LyncSchema = $false
                $LyncVersion = 'Lync Not Installed'
            }
            #endregion OCS/Lync

            $ForestDataProps = @{
                ForestName = $forest.Name
                ForestFunctionalLevel = $forest.ForestMode
                SchemaMaster = $forest.SchemaRoleOwner
                DomainNamingMaster = $forest.NamingRoleOwner
                Sites = @(($forest.Sites | Sort-Object -Property Name | Select-Object Name))
                Domains = @(($forest.Domains | Sort-Object -Property Name | Select-Object Name))
                DomainControllers = $ForestDCs
                DomainControllersCount = $ForestDCs.Count
                GlobalCatalogs = $ForestGCs
                ExchangeServerCount = $ExchangeServerCount
                LyncADContainer = $Lync_ConfigPartition
                LyncServerCount = $LyncServerCount
                LyncPoolCount = $LyncPoolCount
                ExchangeVersion = [string]$ExchangeVersion
                ExchangeServers = $ExchangeServers
                LyncVersion = [string]$LyncVersion
                LyncElements = $Lync_Elements
                TombstoneLifetime = $TombstoneLife
                RecycleBinEnabled = $RecycleBinEnabled
                DeletedObjectLife = $DeletedObjectLife
                LDAPAdminLimits = $LDAPAdminLimits
            }
            $ForestData = New-Object psobject -Property $ForestDataProps
            
            #region AD site subnets
            Write-Verbose -Message ('Get-ADForestReportInformation {0}: Site Subnets - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
            $AD_SiteSubnets = @(Search-AD -Filter '(&(objectClass=subnet))' `
                                       -Properties name,location,siteobject `
                                       -SearchRoot $Path_ADSubnets | 
                                Sort-Object Name)
            Foreach ($Subnet in $AD_SiteSubnets)
            {
                if ($Subnet.siteobject -eq $null)
                {
                    $SiteName = ''
                }
                else
                {
                    $SiteName = Get-ADPathName $Subnet.siteobject -GetElement 0 -ValuesOnly
                }
                #$SiteName = [regex]::Match(($Subnet.siteobject).Split(',')[0], '(?<=CN=).+').Value
                $SiteSubnetProps = @{
                    'Name' = $Subnet.name
                    'Location' = $Subnet.location
                    'SiteName' = $SiteName
                }
                $SiteSubnets += New-Object PSObject -Property $SiteSubnetProps
            }
            #endregion AD site subnets

            #region AD Sites
            Write-Verbose -Message ('Get-ADForestReportInformation {0}: Sites - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
            $Prop_SitesExtended = @('Name',
                                    'DistinguishedName')
            $Prop_SiteConns = @('Name',
                                'DistinguishedName',
                                'Options',
                                'FromServer',
                                'EnabledConnection')
            $AD_SitesExtended = @(Search-AD -Filter '(&(objectClass=site))' `
                                 -Properties $Prop_SitesExtended `
                                 -SearchRoot "LDAP://CN=Sites,$([string]$ConfigNamingContext)")
            $AD_Sites = @([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites)

            ForEach($Site In $AD_Sites) 
            {            
                $SiteDN = [string]($AD_SitesExtended | Where {$_.Name -eq $Site.Name}).DistinguishedName
                $AD_SiteConnections = Search-AD -Filter '(&(objectClass=nTDSConnection))' `
                                                -Properties $Prop_SiteConns `
                                                -SearchRoot "LDAP://$SiteDN"
                $SiteConnections = @()
                if ($AD_SiteConnections -ne $null)
                {
                    Foreach ($SiteConnection in $AD_SiteConnections)
                    {
                        $tmpsiteconn = @($SiteConnection.Options)
                        If(($tmpsiteconn.Count -eq 0) -or ($SiteConnection.Options -eq 0) -or ($Site.Options -eq 'None'))
                        {
                            $SiteConnectionOptions = 'None'
                        }
                        Else
                        {
                            $SiteConnectionOptions = [Enum]::Parse('nTDSSiteConnectionSettingsFlags', $SiteConnection.Options)
                        }
                        
                        $FromServer = Get-ADPathName $SiteConnection.FromServer -GetElement 1 -ValuesOnly
                        $Server = Get-ADPathName $SiteConnection.distinguishedName -GetElement 2 -ValuesOnly

                        $SiteConnProps = @{
                            'DistinguishedName' = $SiteConnection.DistinguishedName
                            'Enabled' = $SiteConnection.EnabledConnection
                            'Options' = $SiteConnectionOptions
                            'FromServer' = $FromServer
                            'Server' = $Server
                        }
                        $SiteConnections += New-Object PSObject -Property $SiteConnProps
                        $AllSiteConnections += New-Object PSObject -Property $SiteConnProps
                    }
                }
                if (($Site.InterSiteTopologyGenerator -ne $null) -and ($Site.InterSiteTopologyGenerator -ne 'None'))
                {
                    $ISTGName = $Site.InterSiteTopologyGenerator | %{[string]$_.Name}
                }
                else
                {
                    $ISTGName = 'None'
                }
                $SiteProps = @{
                        'SiteName' = $Site.Name
                        #'DistinguishedName' = $DistinguishedName
                        'Domains' = @($Site.Domains | %{[string]$_.Name})
                        'Options' = $Site.Options
                        'Location' = $Site.Location
                        'ISTG' = $ISTGName
                        'SiteLinks' = @($Site.SiteLinks | %{[string]$_.Name})
                        'AdjacentSites' = @($Site.AdjacentSites | %{[string]$_.Name})
                        'BridgeheadServers' = ($Site.BridgeheadServers | %{[string]$_.Name})
                        'Connections' = $SiteConnections
                        'ConnectionCount' = $SiteConnections.Count
                        'Subnets' = @($Site.Subnets | %{[string]$_.Name})
                        'SubnetCount' = @($Site.Subnets).Count
                        'Servers' = @($Site.Servers | %{[string]$_.Name})
                        'ServerCount' = @($Site.Servers).Count
                }
                $Sites += New-Object PSObject -Property $SiteProps
            }
            #endregion AD Sites

            #region AD Site Links
            Write-Verbose -Message ('Get-ADForestReportInformation {0}: Site Links - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))

            $AD_SitesLinks = @(Search-AD -Filter '(&(objectClass=siteLink))' `
                                         -Properties cn,replInterval,siteList,options `
                                         -SearchRoot $Path_ADSiteLinks -DontJoinAttributeValues)

            Foreach ($SiteLink in $AD_SitesLinks)
            {
                $SitesInSiteLink = @()
                foreach ($Site in $SiteLink.siteList)
                {
                    $SiteName = Get-ADPathName $Site -GetElement 0 -ValuesOnly
                    $SitesInSiteLink += [string]$SiteName
                }
                $SiteLinkProp = @{
                    Name = $SiteLink.cn
                    repInterval = $SiteLink.replInterval
                    Sites = $SitesInSiteLink
                    ChangeNotification = ($SiteLink.options -eq 1)
                }
                $SiteLinks += new-object psobject -Property $SiteLinkProp
            }
            #endregion AD Site Links

            $SitesSummary = New-Object PSObject -Property @{
                'SiteCount' = $Sites.Count
                'SiteSubnetCount' = $SiteSubnets.Count
                'SiteLinkCount' = $SiteLinks.Count
                'SiteConnectionCount' = $AllSiteConnections.Count
                'SitesWithoutSiteConnections' = @($Sites | Where {$_.ConnectionCount -eq 0}).Count
                'SitesWithoutISTG' = @($Sites | Where {$_.ISTG -eq 'None'}).Count
                'SitesWithoutSubnets' = @($Sites | Where {$_.SubnetCount -eq 0}).Count
                'SitesWithoutServers' = @($Sites | Where {$_.ServerCount -eq 0}).Count
            }
            
            #region Domains
            Write-Verbose -Message ('Get-ADForestReportInformation {0}: Domains - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
            
            ForEach ($Dom in $Forest.Domains) 
            { 
                $CurDomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $Dom.Name)
                $DomainDN = 'dc=' + $Dom.Name.Replace('.', ',dc=')
                $NetBIOSName = Get-NETBiosName $DomainDN $ConfigNamingContext
                if ($Dom.Name -eq ($Forest.RootDomain).Name)
                {
                    $IsForestRoot = $True
                    $SchemaMaster = $forest.SchemaRoleOwner
                    $DomainNamingMaster = $forest.NamingRoleOwner
                }
                else
                {
                    $IsForestRoot = $False
                    $SchemaMaster = 'NA'
                    $DomainNamingMaster = 'NA'
                }
                try
                {
                    $CurDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($CurDomainContext)
                    $CurDomainDetails = [ADSI]"LDAP://$($CurDomain)"
                    $lngMaxPwdAge = $CurDomainDetails.ConvertLargeIntegerToInt64(($CurDomainDetails.maxPwdAge).Value)
                    $lngMinPwdAge = $CurDomainDetails.ConvertLargeIntegerToInt64(($CurDomainDetails.minPwdAge).Value)
                    
                    $DomainFunctionalLevel = $Dom.DomainMode
                    # RID Pool info
                    $Path_RIDManager = "LDAP://CN=RID Manager$,CN=System,$DomainDN"
                    $RIDInfo = Search-AD -Filter '(&(objectClass=rIDManager))' `
                                         -Properties rIDAvailablePool `
                                         -SearchRoot $Path_RIDManager
                    $RIDproperty = $RIDInfo.rIDAvailablePool
                    [int32]$totalSIDS = $($RIDproperty) / ([math]::Pow(2,32))
                    [int64]$temp64val = $totalSIDS * ([math]::Pow(2,32))
                    $RIDsIssued = [int32]($($RIDproperty) - $temp64val)
                    $RIDsRemaining = $totalSIDS - $RIDsIssued
                    $PDCEmulator = $Dom.PdcRoleOwner | Select-Object Name
                    $RIDMaster = $Dom.RidRoleOwner | Select-Object Name
                    $InfrastructureMaster = $Dom.InfrastructureRoleOwner | Select-Object Name
                    $DomainDCs = @($Dom.DomainControllers | Select-Object Name)
                    $lockoutThreshold = $CurDomainDetails.lockoutThreshold
                    $pwdHistoryLength = $CurDomainDetails.pwdHistoryLength
                    $minPwdLength = $CurDomainDetails.minPwdLength
                    $MaxPwdAge = -$lngMaxPwdAge/(600000000 * 1440)
                    $MinPwdAge = -$lngMinPwdAge/(600000000 * 1440)
                    $DomainAccessible = $true
                }
                catch
                {
                    Write-Warning ('Get-ADForestReportInformation: Issue with {0} Domain - {1}' -f $Dom.Name,$_.Exception.Message)
                    $DomainFunctionalLevel = 'NA'
                    $RIDsIssued = 0
                    $RIDsRemaining = 0
                    $PDCEmulator = 'NA'
                    $RIDMaster = 'NA'
                    $InfrastructureMaster = 'NA'
                    $DomainDCs = 'NA'
                    $lockoutThreshold = 0
                    $pwdHistoryLength = 0
                    $minPwdLength = 0
                    $MaxPwdAge = 0
                    $MinPwdAge = 0
                    $DomainAccessible = $false
                }
                $DomainProps = @{
                    DN = $DomainDN
                    Accessible = $DomainAccessible
                    Domain = $Dom.Name
                    NetBIOSName = $NetBIOSName
                    DomainFunctionalLevel = $DomainFunctionalLevel
                    IsForestRoot = $IsForestRoot
                    SchemaMaster = $SchemaMaster
                    DomainNamingMaster = $DomainNamingMaster
                    PDCEmulator = $PDCEmulator
                    RIDMaster = $RIDMaster
                    InfrastructureMaster = $InfrastructureMaster
                    DomainControllers = $DomainDCs
                    lockoutThreshold = $lockoutThreshold
                    pwdHistoryLength = $pwdHistoryLength
                    maxPwdAge = $MaxPwdAge
                    minPwdAge = $MinPwdAge
                    minPwdLength = $minPwdLength
                    RIDSIssued = $RIDsIssued
                    RIDSRemaining = $RIDsRemaining
                    #Sid = $DomSid
                }
                $Domains += New-Object psobject -Property $DomainProps
                if ($DomainAccessible)
                {
                    #region DCs
                    Write-Verbose -Message ('Get-ADForestReportInformation {0}: DCs - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    ForEach ($DC in $Dom.DomainControllers)
                    {
                        $IsGC = $false
                        $IsInfraMaster = $false
                        $IsNamingMaster = $false
                        $IsSchemaMaster = $false
                        $IsRidMaster = $false
                        $IsPdcMaster = $false
                        
                        if ($GCNames -match $DC.Name) { $IsGC = $true }
                        if ($DC.Roles -match 'RidRole') { $IsRidMaster = $true }
                        if ($DC.Roles -match 'PdcRole') { $IsPdcMaster = $true }
                        if ($DC.Roles -match 'InfrastructureRole') { $IsInfraMaster = $true }
                        if ($DC.Roles -match 'SchemaRole') { $IsSchemaMaster = $true }
                        if ($DC.Roles -match 'NamingRole') { $IsNamingMaster = $true }
                        $DCName = [string]$DC.Name
                        $DCName = $DCName.Split('.')[0]
                        $DCProps = @{
                            Forest = ($Dom.Forest).Name
                            Domain = $Dom.Name
                            Site = $DC.SiteName
                            Name = $DCName
                            OS = $DC.OSVersion
                            CurrentTime = $DC.CurrentTime
                            IPAddress = $DC.IPAddress
                          #  HighestUSN = $DC.HighestCommittedUsn
                            IsGC = $IsGC
                            IsInfraMaster = $IsInfraMaster
                            IsNamingMaster = $IsNamingMaster
                            IsSchemaMaster = $IsSchemaMaster
                            IsRidMaster = $IsRidMaster
                            IsPdcMaster = $IsPdcMaster
                        }
                        $DomainControllers += New-Object psobject -Property $DCProps
                    }
                    #endregion DCs
                 
                    #region DFS information
                    $Props_DFSItems = @( 'Name',
                                         'distinguishedName',
                                         'remoteServerName')
                    $Props_DFSGroupTopology = @( 'Name',
                                                 'distinguishedName',
                                                 'msDFSR-ComputerReference')
                    Write-Verbose -Message ('Get-ADForestReportInformation {0}: DFS - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    $DFSDN = "CN=Dfs-Configuration,CN=System,$($DomainDN)"
                    $DFSItems = @(Search-AD -Filter '(&(objectClass=fTDfs))' `
                                            -Properties $Props_DFSItems `
                                            -SearchRoot "LDAP://$DFSDN")
                    foreach ($DFSItem in $DFSItems)
                    {
                        $DomDFSProps = @{
                            Domain = $Dom.Name
                            DN = $DFSItem.distinguishedName
                            Name = $DFSItem.Name
                            RemoteServerName = $DFSItem.remoteServerName -replace ('\*',"")
                        }
                        $DomainDFS += New-Object psobject -Property $DomDFSProps
                    }
                    #endregion DFS information
                    
                    #region DFSR information
                    Write-Verbose -Message ('Get-ADForestReportInformation {0}: DFSR - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    $DFSRDN = "CN=DFSR-GlobalSettings,CN=System,$($DomainDN)"
                    $DFSRGroups = @(Search-AD -Filter '(&(objectClass=msDFSR-ReplicationGroup))' `
                                           -Properties Name,distinguishedName `
                                           -SearchRoot "LDAP://$($DFSRDN)")
                    foreach ($DFSRGroup in $DFSRGroups)
                    {
                        $DFSRGC = @()
                        $DFSRGTop = @()
                        $DFSRGroupContent = @(Search-AD -Filter '(&(objectClass=msDFSR-ContentSet))' `
                                           -Properties Name `
                                           -SearchRoot "LDAP://CN=Content,$($DFSRGroup.distinguishedName)")
                        $DFSRGroupTopology = @(Search-AD -Filter '(&(objectClass=msDFSR-Member))' `
                                           -Properties $Props_DFSGroupTopology `
                                           -SearchRoot "LDAP://CN=Topology,$($DFSRGroup.distinguishedName)")
                        $DFSRGC = @($DFSRGroupContent | %{$_.Name})
                        foreach ($DFSRGroupTopologyItem in $DFSRGroupTopology)
                        {
                            $DFSRServerName = Get-ADPathName $DFSRGroupTopologyItem.'msDFSR-ComputerReference' -GetElement 0 -ValuesOnly
                            $DFSRGTop += [string]$DFSRServerName
                        }
                        $DomDFSRProps = @{
                            Domain = $Dom.Name
                            Name = $DFSRGroup.Name
                            Content = $DFSRGC
                            RemoteServerName = $DFSRGTop
                        }
                        $DomainDFSR += New-Object psobject -Property $DomDFSRProps
                    }
                    #endregion DFSR information

                    #region AD Trusts
                    $ADProps_Trusts = @( 'trusttype',
                                         'trustattributes',
                                         'trustdirection',
                                         'flatname',
                                         'trustpartner',
                                         'whencreated',
                                         'whenchanged')
                    Write-Verbose -Message ('Get-ADForestReportInformation {0}: Trusts - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    $TrustsDN = "CN=System,$($DomainDN)"
                    $AD_Trusts = @(Search-AD -Filter '(&(objectClass=trustedDomain))' `
                                             -SearchRoot "LDAP://$TrustsDN" `
                                             -Properties $ADProps_Trusts)
                    Foreach ($Trust in $AD_Trusts)
                    {
                        switch ($Trust.trusttype) 
                        {
                            1 { $TrustType = 'Downlevel (Windows NT)'}
                            2 { $TrustType = 'Uplevel (Active Directory)'}
                            3 { $TrustType = 'MIT (non-Windows)'}
                            4 { $TrustType = 'DCE (Theoretical)'}
                            default { $TrustType = $Trust.trusttype }
                        }
                        $TrustAttributes = [Enum]::Parse('MSTrustAttributeFlags', $Trust.trustattributes)
                        switch ($Trust.trustdirection)
                        {
                            1 { $TrustDirection = "Inbound"}
                            2 { $TrustDirection = "Outbound"}
                            3 { $TrustDirection = "Bidirectional"}
                            default { $TrustDirection = $Trust.trustdirection }
                        }
                        $TrustProps = @{
                            Domain = $Dom.Name
                            Name = $Trust.flatname
                            TrustedDomain = $Trust.trustpartner
                            Direction = $TrustDirection
                            Attributes = $TrustAttributes
                            TrustType = $TrustType
                            Created = $Trust.whencreated
                            Modified = $Trust.whenchanged
                        }
                        $DomainTrusts += New-Object PSObject -Property $TrustProps
                    }
                    #endregion AD Trusts
                    
                    #region AD Integrated DNS Zones
                    Write-Verbose -Message ('Get-ADForestReportInformation {0}: DNS Zones - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    # Pre-Windows 2003
                    $Path_DNSZoneDN = "LDAP://CN=MicrosoftDNS,CN=System,$DomainDN"
                    $AD_Zones = @(Search-AD -SearchRoot $Path_DNSZoneDN `
                                            -Filter '(objectclass=dnsZone)' `
                                            -Properties name,whencreated,whenchanged,distinguishedName)
                    if ($AD_Zones[0] -ne $null)
                    {
                        Foreach ($DNSZone in $AD_Zones)
                        {
                            $DNSEntryCount = @(Search-AD -SearchRoot "LDAP://$($DNSZone.distinguishedName)" `
                                                         -Filter '(objectclass=dnsNode)')
                            $DNSZoneProps = @{
                                Domain = $Dom.Name
                                AppPartition = 'Legacy'
                                Name  = $DNSZone.name
                                RecordCount = $DNSEntryCount.Count
                                Created = $DNSZone.whencreated
                                Changed = $DNSZone.whenchanged
                            }
                            $DomainDNSZones += New-Object psobject -Property $DNSZoneProps
                        }
                    }
                    $Path_DNSForestZoneDN = "LDAP://DC=ForestDnsZones,$DomainDN"
                    if ([ADSI]::Exists($Path_DNSForestZoneDN))
                    {
                        $AD_ForestZones = @(Search-AD -SearchRoot $Path_DNSForestZoneDN  `
                                                      -Filter '(objectclass=dnsZone)' `
                                                      -Properties name,whencreated,whenchanged,distinguishedName)
                        Foreach ($DNSZone in $AD_ForestZones)
                        {
                            $DNSEntryCount = @(Search-AD -SearchRoot "LDAP://$($DNSZone.distinguishedName)" `
                                                         -Filter '(objectclass=dnsNode)')
                            $DNSZoneProps = @{
                                Domain = $Dom.Name
                                AppPartition = 'Forest'
                                Name  = $DNSZone.name
                                RecordCount = $DNSEntryCount.Count
                                Created = $DNSZone.whencreated
                                Changed = $DNSZone.whenchanged
                            }
                            $DomainDNSZones += New-Object psobject -Property $DNSZoneProps
                        }
                    }
                    
                    $Path_DNSDomainZoneDN = "LDAP://DC=DomainDnsZones,$DomainDN"
                    if ([ADSI]::Exists($Path_DNSDomainZoneDN))
                    {
                        $AD_DomainZones = @(Search-AD -SearchRoot $Path_DNSDomainZoneDN `
                                                      -Filter '(objectclass=dnsZone)' `
                                                      -Properties name,whencreated,whenchanged,distinguishedName)
                        Foreach ($DNSZone in $AD_DomainZones)
                        {
                            $DNSEntryCount = @(Search-AD -SearchRoot "LDAP://$($DNSZone.distinguishedName)" `
                                                         -Filter '(objectclass=dnsNode)')
                            $DNSZoneProps = @{
                                Domain = $Dom.Name
                                AppPartition = 'Domain'
                                Name  = $DNSZone.name
                                RecordCount = $DNSEntryCount.Count
                                Created = $DNSZone.whencreated
                                Changed = $DNSZone.whenchanged
                            }
                            $DomainDNSZones += New-Object psobject -Property $DNSZoneProps
                        }
                    }
                    #endregion AD Integrated DNS Zones
                    
                    #region GPOs
                    Write-Verbose -Message ('Get-ADForestReportInformation {0}: GPOs - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    $AD_DomainGPOs = @(Search-AD -SearchRoot "LDAP://$DomainDN" `
                                                 -Filter '(objectCategory=groupPolicyContainer)' `
                                                 -Properties displayname,whencreated,whenchanged)
                    Foreach ($GPO in $AD_DomainGPOs)
                    {
                        $DomainGPOProps = @{
                            Domain  = $Dom.Name
                            Name    = $GPO.displayname
                            Created = $GPO.whencreated
                            Changed = $GPO.whenchanged
                        }
                        $DomainGPOs += New-Object psobject -Property $DomainGPOProps
                    }
                    #endregion GPOs
                    
                    #region SMS Servers
                    Write-Verbose -Message ('Get-ADForestReportInformation {0}: Domain SMS Servers - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))

                    $SMSServers = @(Search-AD -Filter '(objectclass=mSSMSManagementPoint)' `
                                              -Properties dNSHostName,mSSMSSiteCode,mSSMSVersion,mSSMSDefaultMP,mSSMSDeviceManagementPoint `
                                              -SearchRoot "LDAP://$DomainDN" | 
                                    Select-Object @{n='Domain';e={$Dom.Name}},dNSHostName,mSSMSSiteCode,mSSMSVersion,mSSMSDefaultMP,mSSMSDeviceManagementPoint)
                    #endregion SMS Servers

                    #region SMS Sites
                    Write-Verbose -Message ('Get-ADForestReportInformation {0}: Domain SMS Sites - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    $SMSSites = @()
                    $SMSSiteDetails = @(Search-AD -Filter '(objectclass=mSSMSSite)' `
                                              -Properties Name,mSSMSSiteCode,mSSMSRoamingBoundaries `
                                              -SearchRoot "LDAP://$DomainDN" -DontJoinAttributeValues)
                    $SMSSiteDetails | Foreach {
                        $SMSSiteProps = @{
                            'Domain' = $Dom.Name
                            'Name' = $_.Name
                            'mSSMSSiteCode' = $_.mSSMSSiteCode
                            'mSSMSRoamingBoundaries' = @($_.mSSMSRoamingBoundaries)
                        }
                        $SMSSites += New-Object psobject -Property $SMSSiteProps
                    }
                    #endregion SMS Sites
                    
                    #region NPS Servers
                    Write-Verbose -Message ('Get-ADForestReportInformation {0}: Domain NPS Servers - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    $NPSServers += @((Search-AD -SearchRoot "LDAP://$DomainDN" `
                                              -Filter "(ObjectCategory=group)(Name=RAS and IAS Servers)" `
                                              -Properties member -DontJoinAttributeValues).member | 
                                    Foreach {
                                        [adsi]"LDAP://$($_)" | Select-Object @{n='Domain';e={$Dom.Name}},
                                                                      @{n='Name';e={-join $_.name}},
                                                                      @{n='Type';e={$_.schemaclassname}}
                                    })
                    #endregion NPS Servers
                    
                    #region Printers
                    Write-Verbose -Message ('Get-ADForestReportInformation {0}: Domain Printers - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
                    $DomainPrinters += @(Search-AD -SearchRoot "LDAP://$DomainDN" `
                                              -Filter "(objectCategory=printQueue)" `
                                              -Properties Name,ServerName,printShareName,location,drivername | 
                                        Select-Object @{n='Domain';e={$Dom.Name}},Name,ServerName,printShareName,location,driverName)
                    #endregion Printers
                    #endregion Domains
                }
            }
            
            #region Populate Data
            Write-Verbose -Message ('Get-ADForestReportInformation {0}: Section Data - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
            $SortedRpts | %{ 
                switch ($_.Section) {
                    'ForestSummary' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($ForestData)
                    }
                    'SiteSummary' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($SitesSummary)
                    }
                    'ForestFeatures' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($ForestData)
                    }
                    'ForestDHCPServers' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($DHCPServers)
                    }
                    'ForestExchangeInfo' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($ForestData.ExchangeServers)
                    }
                    'ForestExchangeFederations' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($ExchangeFederations)
                    }
                    'ForestLyncInfo' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($ForestData.LyncElements)
                    }
                    'ForestSiteSummary' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($Sites)
                    }
                    'ForestSiteDetails' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($Sites)
                    }
                    'ForestSiteSubnets' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($SiteSubnets)
                    }
                    'ForestSiteConnections' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($AllSiteConnections)
                    }
                    'ForestSiteLinks' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($SiteLinks)
                    }
                    'ForestDomains' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($Domains)
                    }
                    'ForestDomainDCs' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($DomainControllers)
                    }
                    'ForestDomainPasswordPolicy' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($Domains)
                    }
                    'ForestDomainTrusts' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($DomainTrusts)
                    }
                    'ForestDomainDFSShares' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($DomainDFS)
                    }
                    'ForestDomainDFSRShares' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($DomainDFSR)
                    }
                    'ForestDomainDNSZones' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($DomainDNSZones)
                    }
                    'ForestDomainGPOs' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($DomainGPOs)
                    }
                    'ForestDomainNPSServers' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($NPSServers)
                    }
                    'ForestDomainSCCMServers' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($SMSServers)
                    }
                    'ForestDomainSCCMSites' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($SMSSites)
                    }
                    'ForestDomainPrinters' {
                        $ReportContainer['Sections'][$_]['AllData'][$ForestData.ForestName] = 
                            @($DomainPrinters)
                    }
                }
            }
            #endregion Populate Data
            
            #region Create Diagrams
            Write-Verbose -Message ('Get-ADForestReportInformation {0}: Diagrams - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
            # Replication Connection diagram
            $ReplicationDiagram = @'
digraph test {
 rankdir = LR
 
'@
            ForEach ($Site in $Sites)
            {
                ForEach ($ReplCon in $Site.Connections)
                {
                    $ReplicationDiagram += @"    
            
 "$($ReplCon.FromServer)" -> "$($ReplCon.Server)"[label = "Replicates To"]
"@
               }
            }
            $ReplicationDiagram += @'

}
'@
            If ($AD_CreateDiagramSourceFiles)
            {
                $ReplicationDiagram | Out-File -Encoding ASCII '.\ReplicationDiagram.txt'
            }
            If ($AD_CreateDiagrams)
            {
                
                $ReplicationDiagram | & "$($Graphviz_Path)dot.exe" -Tpng -o ReplicationDiagram.png
            }
            
            # Domain Trust Connection diagram
            $TrustDiagram = @'
digraph test {
 rankdir = LR
 
'@
            ForEach ($Trust in $DomainTrusts)
            {
                $TrustDiagram += @"

 "$($Trust.Domain)" -> "$($Trust.TrustedDomain)"[label = "Trusts"]
"@
            }

            $TrustDiagram += @'

}
'@
            If ($AD_CreateDiagramSourceFiles)
            {
                $TrustDiagram | Out-File -Encoding ASCII '.\DomainTrustDiagram.txt'
            }
            If ($AD_CreateDiagrams)
            {
                $TrustDiagram | & "$($Graphviz_Path)dot.exe" -Tpng -o DomainTrustDiagram.png
            }
            
            # Site Adjacency Diagram
            $SiteAdjacencyDiagram = @'
digraph test {
 rankdir = LR
 
'@
            ForEach ($Site in $Sites)
            {
                Foreach ($AdjSite in $Site.AdjacentSites)
                {
                    $SiteAdjacencyDiagram += @"    
        
     "$($Site.SiteName)" -> "$($AdjSite)"[label = "Adjacent To"]
"@
                }
            }

            $SiteAdjacencyDiagram += @'

}
'@
            If ($AD_CreateDiagramSourceFiles)
            {
                $SiteAdjacencyDiagram | Out-File -Encoding ASCII '.\SiteAdjDiagram.txt'
            }
            If ($AD_CreateDiagrams)
            {
                $SiteAdjacencyDiagram | & "$($Graphviz_Path)dot.exe" -Tpng -o SiteAdjDiagram.png
            }
            #endregion Create Diagrams
            
            $ReportContainer['Configuration']['Assets'] = $ForestData.ForestName
            Return $ForestData.ForestName
            Write-Verbose -Message ('Get-ADForestReportInformation {0}: Finished - {1}' -f $forest.Name,$((New-TimeSpan $verbose_timer ($verbose_timer = get-date)).totalseconds))
        }
    }
}