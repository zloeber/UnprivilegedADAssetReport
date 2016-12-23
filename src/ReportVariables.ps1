# Forest Report comments
$Comment_ForestDomainDCs = 
@'
<tr>
<th class="sectioncolumngrouping" colspan=6>Server Information</th>
<th class="sectioncolumngrouping" colspan=6>Roles</th>
</tr>
'@

# Domain Report comments
$Comment_PrivGroup_EnterpriseAdmins = 
@'
A group that exists only at the forest level of domains. The group is authorized to make forest-wide changes in Active Directory, such as adding child domains. By default, the only member of the group is the Administrator account for the forest root domain.
'@
$Comment_PrivGroup_SchemaAdmins =
@'
A group that exists only at the forest level of domains. The group is authorized to make schema changes in Active Directory. By default, the only member of the group is the Administrator account for the forest root domain. No other accounts should be in this group unless schema upgrades are being done.
'@
$Comment_PrivGroup_DomainAdmins =
@'
Members are authorized to administer the domain. By default, the Domain Admins group is a member of the Administrators group on all computers that have joined a domain, including the domain controllers. Domain Admins is the default owner of any object that is created in the domain's Active Directory by any member of the group. If members of the group create other objects, such as files, the default owner is the Administrators group.
'@
$Comment_PrivGroup_Administrators =
@'
After the initial installation of the operating system, the only member of the group is the Administrator account. When a computer joins a domain, the Domain Admins group is added to the Administrators group. When a server becomes a domain controller, the Enterprise Admins group also is added to the Administrators group. The Administrators group has built-in capabilities that give its members full control over the system. The group is the default owner of any object that is created by a member of the group.
'@
$Comment_PrivGroup_AccountOperators =
@'
Exists only on domain controllers. By default, the group has no members. By default, Account Operators have permission to create, modify, and delete accounts for users, groups, and computers in all containers and organizational units (OUs) of Active Directory except the Builtin container and the Domain Controllers OU. Account Operators do not have permission to modify the Administrators and Domain Admins groups, nor do they have permission to modify the accounts for members of those groups.
'@
$Comment_PrivGroup_ServerOperators =
@'
Exists only on domain controllers. By default, the group has no members. Server Operators can log on to a server interactively; create and delete network shares; start and stop services; back up and restore files; format the hard disk of the computer; and shut down the computer.
'@
$Comment_PrivGroup_BackupOperators =
@'
By default, the group has no members. Backup Operators can back up and restore all files on a computer, regardless of the permissions that protect those files. Backup Operators also can log on to the computer and shut it down.
'@
$Comment_PrivGroup_PrintOperators =
@'
Exists only on domain controllers. By default, the only member is the Domain Users group. Print Operators can manage printers and document queues.
'@
$Comment_PrivGroup_CertPublishers =
@'
Exists only on domain controllers. By default, the only member is the Domain Users group. Print Operators can manage printers and document queues.
'@

# Change this to allow for more or less result properties to span horizontally
#  anything equal to or above this threshold will get displayed vertically instead.
#  (NOTE: This only applies to sections set to be dynamic in html reports)
$HorizontalThreshold = 10

<#
    Configuration
        TOC - Possibly used in the future to create a table of contents
        PreProcessing - Scriptblock to to information gathering
        SkipSectionBreaks - Allows total bypassing of sections of type 
                            'SectionBreak' in reports
        ReportTypes - List all possible report types. The first one listed
                      here will be the default used if none are specified
                      when generating the report.
        Assets - A list of assets which will be reported upon. These are keys in hashes of data
                 broken down by section. In a self contained asset report this will get
                 populated by the PreProcessing information gathering script. Usually
                 this starts out empty and gets automatically filled.
        PostProcessingEnabled - Usually this is true. Currently postprocessing for my scripts
                                rely heavily on a custom function called Format-HTMLTable which,
                                in turn, relies on at least .Net 3.5 sp2 being available for 
                                Linq assemblies. This is done to try and remove the need for
                                custom modules. If you get a bunch of errors about linq not being
                                available you can simply skip post processing by setting this to
                                be false.
#>
$ADForestReport = @{
    'Configuration' = @{
        'TOC'               = $true
        'PreProcessing'     = $ADForestReportPreProcessing
        'SkipSectionBreaks' = $false
        'ReportTypes'       = @('FullDocumentation','ExcelExport')
        'Assets'            = @()
        'PostProcessingEnabled' = $true
    }
    'Sections' = @{
        'Break_ForestInformation' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 0
            'AllData' = @{}
            'Title' = 'Forest Information'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'ExcelExport' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'ForestSummary' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 1
            'AllData' = @{}
            'Title' = 'Forest Summary'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Name';e={$_.ForestName}},
                        @{n='Functional Level';e={$_.ForestFunctionalLevel}},
                        @{n='Domain Naming Master';e={$_.DomainNamingMaster}},
                        @{n='Schema Master';e={$_.SchemaMaster}},
                        @{n='Domain Count';e={($_.Domains).Count}},
                        @{n='DC Server Count';e={$_.DomainControllersCount}},
                        @{n='GC Server Count';e={($_.GlobalCatalogs).Count}},
                        @{n='Exchange Server Count';e={$_.ExchangeServerCount}},
                        @{n='Lync Server Count';e={$_.LyncServerCount}},
                        @{n='Lync Pool Count';e={$_.LyncPoolCount}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Name';e={$_.ForestName}},
                        @{n='Functional Level';e={$_.ForestFunctionalLevel}},
                        @{n='Domain Naming Master';e={$_.DomainNamingMaster}},
                        @{n='Schema Master';e={$_.SchemaMaster}},
                        @{n='Domain Count';e={($_.Domains).Count}},
                        @{n='Site Count';e={($_.Sites).Count}},
                        @{n='DC Server Count';e={$_.DomainControllersCount}},
                        @{n='GC Server Count';e={($_.GlobalCatalogs).Count}},
                        @{n='Exchange Server Count';e={$_.ExchangeServerCount}},
                        @{n='Lync Server Count';e={$_.LyncServerCount}},
                        @{n='Lync Pool Count';e={$_.LyncPoolCount}}
                }
            }
        }
        'SiteSummary' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 2
            'AllData' = @{}
            'Title' = 'Site Summary'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Site Count';e={$_.SiteCount}},
                        @{n='Site Subnet Count';e={$_.SiteSubnetCount}},
                        @{n='Site Link Count';e={$_.SiteLinkCount}},
                        @{n='Site Connection Count';e={$_.SiteConnectionCount}},
                        @{n='Sites Without Site Connections';e={$_.SitesWithotuSiteConnections}},
                        @{n='Sites Without ISTG';e={$_.SitesWithoutISTG}},
                        @{n='Sites Without Subnets';e={$_.SitesWithoutSubnets}},
                        @{n='Sites Without Servers';e={$_.SitesWithoutServers}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Site Count';e={$_.SiteCount}},
                        @{n='Site Subnet Count';e={$_.SiteSubnetCount}},
                        @{n='Site Link Count';e={$_.SiteLinkCount}},
                        @{n='Site Connection Count';e={$_.SiteConnectionCount}},
                        @{n='Sites Without Site Connections';e={$_.SitesWithoutSiteConnections}},
                        @{n='Sites Without ISTG';e={$_.SitesWithoutISTG}},
                        @{n='Sites Without Subnets';e={$_.SitesWithoutSubnets}},
                        @{n='Sites Without Servers';e={$_.SitesWithoutServers}}
                }
            }
        }
        'ForestFeatures' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' =3
            'AllData' = @{}
            'Title' = 'Forest Features'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Recycle Bin Enabled';e={$_.RecycleBinEnabled}},
                        @{n='Tombstone Lifetime';e={$_.TombstoneLifetime}},
                        @{n='Exchange Version';e={$_.ExchangeVersion}},
                        @{n='Lync Version';e={$_.LyncVersion}},
                       # @{n='Deleted Object Lifetime';e={$_.DeletedObjectLife}},
                       # @{n='Total Object Backup Lifetime';e={$_.TotalObjectBackupLife}},
                        @{n='Lync AD Container';e={$_.LyncADContainer}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Recycle Bin Enabled';e={$_.RecycleBinEnabled}},
                        @{n='Tombstone Lifetime';e={$_.TombstoneLifetime}},
                        @{n='Exchange Version';e={$_.ExchangeVersion}},
                        @{n='Lync Version';e={$_.LyncVersion}},
                      #  @{n='Deleted Object Lifetime';e={$_.DeletedObjectLife}},
                      #  @{n='Total Object Backup Lifetime';e={$_.TotalObjectBackupLife}},
                        @{n='Lync AD Container';e={$_.LyncADContainer}}
                }
            }
        }
        'ForestLyncInfo' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 4
            'AllData' = @{}
            'Title' = 'Lync Elements'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Function';e={$_.LyncElement}},
                        @{n='Type';e={$_.LyncElementType}},
                        @{n='FQDN';e={$_.LyncElementFQDN}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Function';e={$_.LyncElement}},
                        @{n='Type';e={$_.LyncElementType}},
                        @{n='FQDN';e={$_.LyncElementFQDN}}
                }
            }
            'PostProcessing' = $LyncElements_Postprocessing
        }
        'ForestExchangeInfo' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 5
            'AllData' = @{}
            'Title' = 'Forest Exchange Servers'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Org';e={$_.Organization}},
                        @{n='Admin Group';e={$_.AdminGroup}},
                        @{n='Name';e={$_.Name}},
                        @{n='Roles';e={$_.Role}},
                        @{n='Site';e={$_.Site}},
                        #@{n='Created';e={$_.Created}},
                        @{n='Serial';e={$_.Serial}},
                        @{n='Product ID';e={$_.ProductID}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Org';e={$_.Organization}},
                        @{n='Admin Group';e={$_.AdminGroup}},
                        @{n='Name';e={$_.Name}},
                        @{n='Roles';e={$_.Role}},
                        @{n='Site';e={$_.Site}},
                        #@{n='Created';e={$_.Created}},
                        @{n='Serial';e={$_.Serial}},
                        @{n='Product ID';e={$_.ProductID}}
                }
            }
        }
        'ForestExchangeFederations' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 6
            'AllData' = @{}
            'Title' = 'Forest Exchange Federations'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Org';e={$_.Organization}},
                        @{n='Name';e={$_.Name}},
                        @{n='Enabled';e={$_.Enabled}},
                        @{n='Domains';e={[string]$_.Domains -replace ' ', "`n`r"}},
                        @{n='Allowed Actions';e={[string]$_.AllowedActions -replace ' ', "`n`r"}},
                        @{n='App URI';e={$_.TargetAppURI}},
                        @{n='Autodiscover EPR';e={$_.TargetAutodiscoverEPR}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Org';e={$_.Organization}},
                        @{n='Name';e={$_.Name}},
                        @{n='Enabled';e={$_.Enabled}},
                        @{n='Domains';e={[string]$_.Domains -replace ' ', "<br />`n`r"}},
                        @{n='Allowed Actions';e={[string]$_.AllowedActions -replace ' ', "<br />`n`r"}}
                        #@{n='App URI';e={$_.TargetAppURI}},
                        #@{n='Autodiscover EPR';e={$_.TargetAutodiscoverEPR}}
                }
            }
        }
        'ForestDHCPServers' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 7
            'AllData' = @{}
            'Title' = 'Registered DHCP Servers'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Created';e={$_.WhenCreated}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Created';e={$_.WhenCreated}}
                }
            }
        }
        'ForestDomainNPSServers' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 8
            'AllData' = @{}
            'Title' = 'Registered NPS Servers'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='Type';e={$_.Type}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='Type';e={$_.Type}}
                }
            }
        }
        'Break_SiteInformation' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 10
            'AllData' = @{}
            'Title' = 'Site Information'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'ExcelExport' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'ForestSiteSummary' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 11
            'AllData' = @{}
            'Title' = 'Site Summary'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.SiteName}},
                        @{n='Location';e={$_.Location}},
                        @{n='Domains';e={[string]$_.Domains -replace ' ', "`n`r"}},
                        @{n='DCs';e={[string]$_.Servers -replace ' ', "`n`r"}},
                        @{n='Subnets';e={[string]$_.Subnets  -replace ' ', "`n`r"}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.SiteName}},
                        @{n='Location';e={$_.Location}},
                        @{n='Domains';e={[string]$_.Domains -replace ' ', "<br />`n`r"}},
                        @{n='DCs';e={[string]$_.Servers -replace ' ', "<br />`n`r"}},
                        @{n='Subnets';e={[string]$_.Subnets  -replace ' ', "<br />`n`r"}}
                }
            }
        }
        'ForestSiteDetails' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 12
            'AllData' = @{}
            'Title' = 'Site Details'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.SiteName}},
                        @{n='Options';e={$_.Options}},
                        @{n='ISTG';e={$_.ISTG}},
                        @{n='SiteLinks';e={[string]$_.SiteLinks -replace ' ', "`n`r"}},
                        @{n='BridgeheadServers';e={[string]$_.BridgeheadServers -replace ' ', "`n`r"}},
                        @{n='AdjacentSites';e={[string]$_.AdjacentSites -replace ' ', "`n`r"}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.SiteName}},
                        @{n='Options';e={$_.Options}},
                        @{n='ISTG';e={$_.ISTG}},
                        @{n='SiteLinks';e={[string]$_.SiteLinks -replace ' ', "<br />`n`r"}},
                        @{n='BridgeheadServers';e={[string]$_.BridgeheadServers -replace ' ', "<br />`n`r"}},
                        @{n='AdjacentSites';e={[string]$_.AdjacentSites -replace ' ', "<br />`n`r"}}
                }
            }
        }
        'ForestSiteSubnets' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 13
            'AllData' = @{}
            'Title' = 'Site Subnets'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Subnet';e={$_.Name}},
                        @{n='Site Name';e={$_.SiteName}},
                        @{n='Location';e={$_.Location}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Subnet';e={$_.Name}},
                        @{n='Site Name';e={$_.SiteName}},
                        @{n='Location';e={$_.Location}}
                }
            }
        }
        'ForestSiteConnections' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 14
            'AllData' = @{}
            'Title' = 'Site Connections'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Enabled';e={$_.Enabled}},
                        @{n='Options';e={$_.Options}},
                        @{n='From';e={$_.FromServer}},
                        @{n='To';e={$_.Server}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Enabled';e={$_.Enabled}},
                        @{n='Options';e={$_.Options}},
                        @{n='From';e={$_.FromServer}},
                        @{n='To';e={$_.Server}}
                }
            }
            'PostProcessing' = $ForestSiteConnections_Postprocessing
        }
        'ForestSiteLinks' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 15
            'AllData' = @{}
            'Title' = 'Site Links'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Replication Interval';e={$_.repInterval}},
                        @{n='Sites';e={[string]$_.Sites -replace ' ', "`n`r"}},
                        @{n='Change Notification Enabled';e={$_.ChangeNotification}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Replication Interval';e={$_.repInterval}},
                        @{n='Sites';e={[string]$_.Sites -replace ' ', "<br />`n`r"}},
                        @{n='Change Notification Enabled';e={$_.ChangeNotification}}
                }
            }
            'PostProcessing' = $False
        }
        'Break_DomainInformation' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 20
            'AllData' = @{}
            'Title' = 'Domain Information'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'ExcelExport' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'ForestDomains' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 21
            'AllData' = @{}
            'Title' = 'Domains'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Domain}},
                        @{n='NetBIOS';e={$_.NetBIOSName}},
                        @{n='Functional Level';e={$_.DomainFunctionalLevel}},
                        @{n='Forest Root';e={$_.IsForestRoot}},
                        @{n='RIDs Issued';e={$_.RIDsIssued}},
                        @{n='RIDs Remaining';e={$_.RIDsRemaining}},
                        @{n='Naming Master';e={$_.DomainNamingMaster}},
                        @{n='Schema Master';e={$_.SchemaMaster}},
                        @{n='PDC Emulator';e={$_.PDCEmulator}},
                        @{n='RID Master';e={$_.RIDMaster}},
                        @{n='Infra Master';e={$_.InfrastructureMaster}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Domain}},
                        @{n='NetBIOS';e={$_.NetBIOSName}},
                        @{n='Functional Level';e={$_.DomainFunctionalLevel}},
                        @{n='Forest Root';e={$_.IsForestRoot}},
                        @{n='RIDs Issued';e={$_.RIDsIssued}},
                        @{n='RIDs Remaining';e={$_.RIDsRemaining}}
                }
            }
        }
        'ForestDomainPasswordPolicy' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 22
            'AllData' = @{}
            'Title' = 'Domain Password Policies'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Domain}},
                        @{n='NetBIOS';e={$_.NetBIOSName}},
                        @{n='Lockout Threshold';e={$_.lockoutThreshold}},
                        @{n='Password History Length';e={$_.pwdHistoryLength}},
                        @{n='Max Password Age';e={$_.maxPwdAge}},
                        @{n='Min Password Age';e={$_.minPwdAge}},
                        @{n='Min Password Length';e={$_.minPwdLength}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Domain}},
                        @{n='NetBIOS';e={$_.NetBIOSName}},
                        @{n='Lockout Threshold';e={$_.lockoutThreshold}},
                        @{n='Password History Length';e={$_.pwdHistoryLength}},
                        @{n='Max Password Age';e={$_.maxPwdAge}},
                        @{n='Min Password Age';e={$_.minPwdAge}},
                        @{n='Min Password Length';e={$_.minPwdLength}}
                }
            }
        }
        'ForestDomainDCs' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 23
            'AllData' = @{}
            'Title' = 'Domain Controllers'
            'Type' = 'Section'
            'Comment' = $Comment_ForestDomainDCs
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Site';e={$_.Site}},
                        @{n='Name';e={$_.Name}},
                        @{n='OS';e={$_.OS}},
                        @{n='Time';e={$_.CurrentTime}},
                        @{n='IP';e={$_.IPAddress}},
                        @{n='GC';e={$_.IsGC}},
                        @{n='Infra';e={$_.IsInfraMaster}},
                        @{n='Naming';e={$_.IsNamingMaster}},
                        @{n='Schema';e={$_.IsSchemaMaster}},
                        @{n='RID';e={$_.IsRidMaster}},
                        @{n='PDC';e={$_.IsPdcMaster}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Site';e={$_.Site}},
                        @{n='Name';e={$_.Name}},
                        @{n='OS';e={$_.OS}},
                        @{n='Time';e={$_.CurrentTime}},
                        @{n='IP';e={$_.IPAddress}},
                        @{n='GC';e={$_.IsGC}},
                        @{n='Infra';e={$_.IsInfraMaster}},
                        @{n='Naming';e={$_.IsNamingMaster}},
                        @{n='Schema';e={$_.IsSchemaMaster}},
                        @{n='RID';e={$_.IsRidMaster}},
                        @{n='PDC';e={$_.IsPdcMaster}}
                }
            }
            'PostProcessing' = $ForestDomainDCs_Postprocessing
        }
        'ForestDomainTrusts' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 24
            'AllData' = @{}
            'Title' = 'Domain Trusts'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Trusted Domain';e={$_.TrustedDomain}},
                        @{n='Direction';e={$_.Direction}},
                        @{n='Attributes';e={$_.Attributes}},
                        @{n='Trust Type';e={$_.TrustType}},
                        @{n='Created';e={$_.Created}},
                        @{n='Modified';e={$_.Modified}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Trusted Domain';e={$_.TrustedDomain}},
                        @{n='Direction';e={$_.Direction}},
                        @{n='Attributes';e={$_.Attributes}},
                        @{n='Trust Type';e={$_.TrustType}},
                        @{n='Created';e={$_.Created}},
                        @{n='Modified';e={$_.Modified}}
                }
            }
        }
        'ForestDomainDFSShares' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 25
            'AllData' = @{}
            'Title' = 'Domain DFS Shares'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='DN';e={$_.DN}},
                        @{n='Remote Server';e={$_.RemoteServerName}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='DN';e={$_.DN}},
                        @{n='Remote Server';e={$_.RemoteServerName}}
                }
            }
        }
        'ForestDomainDFSRShares' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 26
            'AllData' = @{}
            'Title' = 'Domain DFSR Shares'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='Content';e={[string]$_.Content -replace ' ', "`n`r"}},
                        @{n='Remote Servers';e={[string]$_.RemoteServerName -replace ' ', "`n`r"}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='Content';e={[string]$_.Content -replace ' ', "<br />`n`r"}},
                        @{n='Remote Servers';e={[string]$_.RemoteServerName -replace ' ', "<br />`n`r"}}
                }
            }
        }
        'ForestDomainDNSZones' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 27
            'AllData' = @{}
            'Title' = 'Domain Integrated DNS Zones'
            'Type' = 'Section'
            'Comment' = 'Active Directory integrated DNS zones. Zone names containing CNF: or InProgress may be duplicate and should be reviewed.'
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Partition';e={$_.AppPartition}},
                        @{n='Name';e={$_.Name}},
                        @{n='Record Count';e={$_.RecordCount}},
                        @{n='Created';e={$_.Created}},
                        @{n='Changed';e={$_.Changed}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Partition';e={$_.AppPartition}},
                        @{n='Name';e={$_.Name}},
                        @{n='Record Count';e={$_.RecordCount}},
                        @{n='Created';e={$_.Created}},
                        @{n='Changed';e={$_.Changed}}
                }
            }
            'PostProcessing' = $ForestDomainDNSZones_Postprocessing
        }
        'ForestDomainGPOs' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 28
            'AllData' = @{}
            'Title' = 'Domain GPOs'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='Created';e={$_.Created}},
                        @{n='Changed';e={$_.Changed}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='Created';e={$_.Created}},
                        @{n='Changed';e={$_.Changed}}
                }
            }
        }
        'ForestDomainPrinters' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 30
            'AllData' = @{}
            'Title' = 'Domain Registered Printers'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='ServerName';e={$_.serverName}},
                        @{n='ShareName';e={$_.printShareName}},
                        @{n='Location';e={$_.location}},
                        @{n='DriverName';e={$_.driverName}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='ServerName';e={$_.serverName}},
                        @{n='ShareName';e={$_.printShareName}},
                        @{n='Location';e={$_.location}},
                        @{n='DriverName';e={$_.driverName}}
                }
            }
        }
        'ForestDomainSCCMServers' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 31
            'AllData' = @{}
            'Title' = 'Registered SCCM Servers'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.dNSHostName}},
                        @{n='Site Code';e={$_.mSSMSSiteCode}},
                        @{n='Version';e={$_.mSSMSVersion}},
                        @{n='Default MP';e={$_.mSSMSDefaultMP}},
                        @{n='Device MP';e={$_.mSSMSDeviceManagementPoint}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.dNSHostName}},
                        @{n='Site Code';e={$_.mSSMSSiteCode}},
                        @{n='Version';e={$_.mSSMSVersion}},
                        @{n='Default MP';e={$_.mSSMSDefaultMP}},
                        @{n='Device MP';e={$_.mSSMSDeviceManagementPoint}}
                }
            }
        }
        'ForestDomainSCCMSites' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 32
            'AllData' = @{}
            'Title' = 'Registered SCCM Sites'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'ExcelExport' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='Site Code';e={$_.mSSMSSiteCode}},
                        @{n='Roaming Boundries';e={[string]$_.mSSMSRoamingBoundaries -replace ' ', "`n`r"}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Domain';e={$_.Domain}},
                        @{n='Name';e={$_.Name}},
                        @{n='Site Code';e={$_.mSSMSSiteCode}},
                        @{n='Roaming Boundries';e={[string]$_.mSSMSRoamingBoundaries -replace ' ', "<br />`n`r"}}
                }
            }
        }
    }
} 

$ADDomainReport = @{
    'Configuration' = @{
        'TOC'                   = $true
        'PreProcessing'         = $ADDomainReportPreProcessing
        'SkipSectionBreaks'     = $false
        'ReportTypes'           = @('FullDocumentation')
        'Assets'                = @()
        'PostProcessingEnabled' = $true
    }
    'Sections' = @{
        'Break_Stats' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 0
            'AllData' = @{}
            'Title' = 'Domain Statistics'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'UserAccountStats1' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 1
            'AllData' = @{}
            'Title' = 'User Account Statistics'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Total User Accounts';e={$_.Total}},
                        @{n='Enabled';e={$_.Enabled}},
                        @{n='Disabled';e={$_.Disabled}},
                        @{n='Locked';e={$_.Locked}},
                        @{n='Password Does Not Expire';e={$_.PwdDoesNotExpire}},
                        @{n='Password Must Change';e={$_.PwdMustChange}}
                }
            }
        }
        'UserAccountStats2' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 2
            'AllData' = @{}
            'Title' = 'User Account Statistics'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Password Not Required';e={$_.PwdNotRequired}},
                        @{n='Dial-in Enabled';e={$_.DialInEnabled}},
                        @{n='Control Access With NPS';e={$_.ControlAccessWithNPS}},
                        @{n='Unconstrained Delegation';e={$_.UnconstrainedDelegation}},
                        @{n='Not Trusted For Delegation';e={$_.NotTrustedForDelegation}},
                        @{n='No Pre-Auth Required';e={$_.NoPreAuthRequired}}
                }
            }
        }
        'GroupStats' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 3
            'AllData' = @{}
            'Title' = 'Group Statistics'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Total Groups';e={$_.Total}},
                        @{n='Built-in';e={$_.Builtin}},
                        @{n='Universal Security';e={$_.UniversalSecurity}},
                        @{n='Universal Distribution';e={$_.UniversalDist}},
                        @{n='Global Security';e={$_.GlobalSecurity}},
                        @{n='Global Distribution';e={$_.GlobalDist}},
                        @{n='Domain Local Security';e={$_.DomainLocalSecurity}},
                        @{n='Domain Local Distribution';e={$_.DomainLocalDist}}
                }
            }
        }
        'PrivGroupStats' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 10
            'AllData' = @{}
            'Title' = 'Privileged Group Statistics'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Default Name';e={$_.AdminGroup}},
                        @{n='Current Name';e={$_.DisplayName}},
                        @{n='Member Count';e={$_.MemberCount}}
                }
            }
        }
        'PrivGroup_EnterpriseAdmins' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 20
            'AllData' = @{}
            'Title' = 'Enterprise Administrators'
            'Type' = 'Section'
            'Comment' = $Comment_PrivGroup_EnterpriseAdmins
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Logon ID';e={$_.sAMAccountName}},
                        @{n='Name';e={$_.name}},
                        @{n='Pwd Age (Days)';e={$_.PasswordAge}},
                        @{n='Last Logged In';e={$_.lastlogontimestamp}},
                        @{n='No Pwd Expiry';e={$_.DONT_EXPIRE_PASSWD}},
                        @{n='Pwd Reversable';e={$_.ENCRYPTED_TEXT_PASSWORD_ALLOWED}},
                        @{n='Pwd Not Req.';e={$_.PASSWD_NOTREQD}}
                }
            }
            'PostProcessing' = $ADPrivUser_Postprocessing
        }
        'PrivGroup_SchemaAdmins' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 21
            'AllData' = @{}
            'Title' = 'Schema Administrators'
            'Type' = 'Section'
            'Comment' = $Comment_PrivGroup_SchemaAdmins
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Logon ID';e={$_.sAMAccountName}},
                        @{n='Name';e={$_.name}},
                        @{n='Pwd Age (Days)';e={$_.PasswordAge}},
                        @{n='Last Logged In';e={$_.lastlogontimestamp}},
                        @{n='No Pwd Expiry';e={$_.DONT_EXPIRE_PASSWD}},
                        @{n='Pwd Reversable';e={$_.ENCRYPTED_TEXT_PASSWORD_ALLOWED}},
                        @{n='Pwd Not Req.';e={$_.PASSWD_NOTREQD}}
                }
            }
            'PostProcessing' = $ADPrivUser_Postprocessing
        }
        'PrivGroup_DomainAdmins' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 22
            'AllData' = @{}
            'Title' = 'Domain Administrators'
            'Type' = 'Section'
            'Comment' = $Comment_PrivGroup_DomainAdmins
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Logon ID';e={$_.sAMAccountName}},
                        @{n='Name';e={$_.name}},
                        @{n='Pwd Age (Days)';e={$_.PasswordAge}},
                        @{n='Last Logged In';e={$_.lastlogontimestamp}},
                        @{n='No Pwd Expiry';e={$_.DONT_EXPIRE_PASSWD}},
                        @{n='Pwd Reversable';e={$_.ENCRYPTED_TEXT_PASSWORD_ALLOWED}},
                        @{n='Pwd Not Req.';e={$_.PASSWD_NOTREQD}}
                }
            }
            'PostProcessing' = $ADPrivUser_Postprocessing
        }
        'PrivGroup_Administrators' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 23
            'AllData' = @{}
            'Title' = 'Administrators'
            'Type' = 'Section'
            'Comment' = $Comment_PrivGroup_Administrators
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Logon ID';e={$_.sAMAccountName}},
                        @{n='Name';e={$_.name}},
                        @{n='Pwd Age (Days)';e={$_.PasswordAge}},
                        @{n='Last Logged In';e={$_.lastlogontimestamp}},
                        @{n='No Pwd Expiry';e={$_.DONT_EXPIRE_PASSWD}},
                        @{n='Pwd Reversable';e={$_.ENCRYPTED_TEXT_PASSWORD_ALLOWED}},
                        @{n='Pwd Not Req.';e={$_.PASSWD_NOTREQD}}
                }
            }
            'PostProcessing' = $ADPrivUser_Postprocessing
        }
        'PrivGroup_ServerOperators' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 24
            'AllData' = @{}
            'Title' = 'Server Operators'
            'Type' = 'Section'
            'Comment' = $Comment_PrivGroup_ServerOperators
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Logon ID';e={$_.sAMAccountName}},
                        @{n='Name';e={$_.name}},
                        @{n='Pwd Age (Days)';e={$_.PasswordAge}},
                        @{n='Last Logged In';e={$_.lastlogontimestamp}},
                        @{n='No Pwd Expiry';e={$_.DONT_EXPIRE_PASSWD}},
                        @{n='Pwd Reversable';e={$_.ENCRYPTED_TEXT_PASSWORD_ALLOWED}},
                        @{n='Pwd Not Req.';e={$_.PASSWD_NOTREQD}}
                }
            }
            'PostProcessing' = $ADPrivUser_Postprocessing
        }
        'PrivGroup_BackupOperators' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 25
            'AllData' = @{}
            'Title' = 'Backup Operators'
            'Type' = 'Section'
            'Comment' = $Comment_PrivGroup_BackupOperators
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Logon ID';e={$_.sAMAccountName}},
                        @{n='Name';e={$_.name}},
                        @{n='Pwd Age (Days)';e={$_.PasswordAge}},
                        @{n='Last Logged In';e={$_.lastlogontimestamp}},
                        @{n='No Pwd Expiry';e={$_.DONT_EXPIRE_PASSWD}},
                        @{n='Pwd Reversable';e={$_.ENCRYPTED_TEXT_PASSWORD_ALLOWED}},
                        @{n='Pwd Not Req.';e={$_.PASSWD_NOTREQD}}
                }
            }
            'PostProcessing' = $ADPrivUser_Postprocessing
        }
        'PrivGroup_AccountOperators' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 26
            'AllData' = @{}
            'Title' = 'Account Operators'
            'Type' = 'Section'
            'Comment' = $Comment_PrivGroup_AccountOperators
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Logon ID';e={$_.sAMAccountName}},
                        @{n='Name';e={$_.name}},
                        @{n='Pwd Age (Days)';e={$_.PasswordAge}},
                        @{n='Last Logged In';e={$_.lastlogontimestamp}},
                        @{n='No Pwd Expiry';e={$_.DONT_EXPIRE_PASSWD}},
                        @{n='Pwd Reversable';e={$_.ENCRYPTED_TEXT_PASSWORD_ALLOWED}},
                        @{n='Pwd Not Req.';e={$_.PASSWD_NOTREQD}}
                }
            }
            'PostProcessing' = $ADPrivUser_Postprocessing
        }
        'PrivGroup_CertPublishers' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 27
            'AllData' = @{}
            'Title' = 'Certificate Publishers'
            'Type' = 'Section'
            'Comment' = $Comment_PrivGroup_CertPublishers
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Logon ID';e={$_.sAMAccountName}},
                        @{n='Name';e={$_.name}},
                        @{n='Pwd Age (Days)';e={$_.PasswordAge}},
                        @{n='Last Logged In';e={$_.lastlogontimestamp}},
                        @{n='No Pwd Expiry';e={$_.DONT_EXPIRE_PASSWD}},
                        @{n='Pwd Reversable';e={$_.ENCRYPTED_TEXT_PASSWORD_ALLOWED}},
                        @{n='Pwd Not Req.';e={$_.PASSWD_NOTREQD}}
                }
            }
            'PostProcessing' = $ADPrivUser_Postprocessing
        }
        'PrivGroup_PrintOperators' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 28
            'AllData' = @{}
            'Title' = 'Print Operators'
            'Type' = 'Section'
            'Comment' = $Comment_PrivGroup_PrintOperators
            'ReportTypes' = @{
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Logon ID';e={$_.sAMAccountName}},
                        @{n='Name';e={$_.name}},
                        @{n='Pwd Age (Days)';e={$_.PasswordAge}},
                        @{n='Last Logged In';e={$_.lastlogontimestamp}},
                        @{n='No Pwd Expiry';e={$_.DONT_EXPIRE_PASSWD}},
                        @{n='Pwd Reversable';e={$_.ENCRYPTED_TEXT_PASSWORD_ALLOWED}},
                        @{n='Pwd Not Req.';e={$_.PASSWD_NOTREQD}}
                }
            }
            'PostProcessing' = $ADPrivUser_Postprocessing
        }
    }
}