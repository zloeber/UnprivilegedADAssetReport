$IsPS5 = ($PSVersionTable.PSVersion).Major -ge 5

if ($IsPS5) {
    Write-Verbose "Powershell version 5 detected, using builtin Flags instead of add-type definitions."
    [Flags()] enum nTDSSiteConnectionSettingsFlags {
        IS_GENERATED = 1
        TWOWAY_SYNC = 2
        OVERRIDE_NOTIFY_DEFAULT = 4
        USE_NOTIFY = 8
        DISABLE_INTERSITE_COMPRESSION = 10
        OPT_USER_OWNED_SCHEDULE = 20
    }
    [Flags()] enum MSExchCurrentServerRolesFlags {
        NONE = 1
        MAILBOX = 2
        CLIENT_ACCESS = 4
        UM = 10
        HUB_TRANSPORT  = 20
        EDGE_TRANSPORT = 40  
    }
    [Flags()] enum nTDSSiteSettingsFlags {
        IS_AUTO_TOPOLOGY_DISABLED = 1
        IS_TOPL_CLEANUP_DISABLED = 2
        IS_TOPL_MIN_HOPS_DISABLED = 4
        IS_TOPL_DETECT_STALE_DISABLED = 8
        IS_INTER_SITE_AUTO_TOPOLOGY_DISABLED = 10
        IS_GROUP_CACHING_ENABLED = 20
        FORCE_KCC_WHISTLER_BEHAVIOR = 40
        FORCE_KCC_W2K_ELECTION = 80
        IS_RAND_BH_SELECTION_DISABLED = 100
        IS_SCHEDULE_HASHING_ENABLED = 200
        IS_REDUNDANT_SERVER_TOPOLOGY_ENABLED = 400
    }
    [Flags()] enum MSTrustAttributeFlags {
        NON_TRANSITIVE = 1
        UPLEVEL_ONLY = 2
        QUARANTINED_DOMAIN = 4
        FOREST_TRANSITIVE = 8
        CROSS_ORGANIZATION = 10
        WITHIN_FOREST = 20
        TREAT_AS_EXTERNAL  = 40
        USES_RC4_ENCRYPTION = 80
    }
}

else {
    Add-Type -TypeDefinition @" 
        [System.Flags]
        public enum nTDSSiteConnectionSettingsFlags {
            IS_GENERATED                  = 0x00000001,
            TWOWAY_SYNC                   = 0x00000002,
            OVERRIDE_NOTIFY_DEFAULT       = 0x00000004,
            USE_NOTIFY                    = 0x00000008,
            DISABLE_INTERSITE_COMPRESSION = 0x00000010,
            OPT_USER_OWNED_SCHEDULE       = 0x00000020  
        }
        [System.Flags]
        public enum MSExchCurrentServerRolesFlags {
            NONE           = 0x00000001,
            MAILBOX        = 0x00000002,
            CLIENT_ACCESS  = 0x00000004,
            UM             = 0x00000010,
            HUB_TRANSPORT  = 0x00000020,
            EDGE_TRANSPORT = 0x00000040  
        }
        [System.Flags]
        public enum nTDSSiteSettingsFlags {
            IS_AUTO_TOPOLOGY_DISABLED            = 0x00000001,
            IS_TOPL_CLEANUP_DISABLED             = 0x00000002,
            IS_TOPL_MIN_HOPS_DISABLED            = 0x00000004,
            IS_TOPL_DETECT_STALE_DISABLED        = 0x00000008,
            IS_INTER_SITE_AUTO_TOPOLOGY_DISABLED = 0x00000010,
            IS_GROUP_CACHING_ENABLED             = 0x00000020,
            FORCE_KCC_WHISTLER_BEHAVIOR          = 0x00000040,
            FORCE_KCC_W2K_ELECTION               = 0x00000080,
            IS_RAND_BH_SELECTION_DISABLED        = 0x00000100,
            IS_SCHEDULE_HASHING_ENABLED          = 0x00000200,
            IS_REDUNDANT_SERVER_TOPOLOGY_ENABLED = 0x00000400
        }
        [System.Flags]
        public enum MSTrustAttributeFlags {
            NON_TRANSITIVE      = 0x00000001,
            UPLEVEL_ONLY        = 0x00000002,
            QUARANTINED_DOMAIN  = 0x00000004,
            FOREST_TRANSITIVE   = 0x00000008,
            CROSS_ORGANIZATION  = 0x00000010,
            WITHIN_FOREST       = 0x00000020,
            TREAT_AS_EXTERNAL   = 0x00000040,
            USES_RC4_ENCRYPTION = 0x00000080
        }
"@
}

# If you are color coding the domain reports this will control password age colorization
$AD_PwdAgeWarn = 60
$AD_PwdAgeAlert = 90
$AD_PwdAgeHealthy = 60

# A list of user attributes to normalize across all users.
# When an attribute doesn't exist (a non-mailbox enabled
# account for instance), it will be added with a $null value.
# These will all be exported if $EXPORTTOCSV_USERS is $true
$UserAttribs = @(
    'cn',
    'displayName',
    'givenName',
    'sn',
    'name',
    'sAMAccountName',
    'sAMAccountType',
    'whenChanged',
    'whenCreated',
    'pwdLastSet',
    'admincount',
    'accountExpires',
    'badPasswordTime',
    'badPwdCount',
    'lastLogon',
    'lastLogoff',
    'logonCount',
    'useraccountcontrol',
    'lastlogontimestamp',
    'homeMDB',
    'homeMTA',
    'mail',
    'proxyAddresses',
    'mailNickname',
    #'legacyExchangeDN',
    #'showInAddressBook',
    'msexchalobjectversion',
    #'msexchdelegatelistbl',        # Could be interesting for a seperate report
    'msexchhomeservername',
    'msexchrecipientdisplaytype',
    'msexchrecipienttypedetails',
    'msexchumdtmfmap',
    'msexchuseraccountcontrol',
    'msexchuserculture',
    'msexchversion',
    'msexchwhenmailboxcreated',
    'msnpallowdialin',
    'msRTCSIP-PrimaryHomeServer',
    'msRTCSIP-PrimaryUserAddress',
    'msRTCSIP-UserEnabled',
    'msRTCSIP-Line',
    'msRTCSIP-FederationEnabled',
    'msRTCSIP-InternetAccessEnabled'
)

# These are what we will attempt to report upon later on as 'privileged' groups
$AD_PrivilegedGroups = @(
    'Enterprise Admins',
    'Schema Admins',
    'Domain Admins',
    'Administrators',
    'Cert Publishers',
    'Account Operators',
    'Server Operators',
    'Backup Operators',
    'Print Operators'
)

$Attrib_User_MSExchangeVersion = @{
    # $null = Exchange 2003 and earlier
    '4535486012416' = '2007'
    '44220983382016' = '2010'
}

#Schema constants
$SchemaHashExchange = 
@{
    4397='Exchange Server 2000 RTM'
    4406='Exchange Server 2000 SP3'
    6870='Exchange Server 2003 RTM'
    6936='Exchange Server 2003 SP3'
    10628='Exchange Server 2007 RTM'
    10637='Exchange Server 2007 RTM'
    11116='Exchange 2007 SP1'
    14622='Exchange 2007 SP2 or Exchange 2010 RTM'
    14625='Exchange 2007 SP3'
    14726='Exchange 2010 SP1'
    14732='Exchange 2010 SP2'
    14734='Exchange 2010 SP3'
    15137='Exchange 2013 RTM'
    15254='Exchange 2013 CU1'
    15281='Exchange 2013 CU2'
    15283='Exchange 2013 CU3'
}
$SchemaHashLync = 
@{
    1006="LCS 2005"
    1007="OCS 2007 R1"
    1008="OCS 2007 R2"
    1100="Lync Server 2010"
    1150="Lync Server 2013"
}

# AD DC capabilities list (http://www.ldapexplorer.com/en/manual/103010700-connection-rootdse.htm)
# - Primarily used to determine if a DC is RODC or not (Const LDAP_CAP_ACTIVE_DIRECTORY_PARTIAL_SECRETS_OID = "1.2.840.113556.1.4.1920")
$AD_Capabilities = @{
    '1.2.840.113556.1.4.319' = 'Paged results'
    '1.2.840.113556.1.4.417' = 'Show deleted objects'
    '1.2.840.113556.1.4.473' = 'Sort results'
    '1.2.840.113556.1.4.474' = 'Sort results response'
    '1.2.840.113556.1.4.521' = 'Cross domain move'
    '1.2.840.113556.1.4.528' = 'Server notification'
    '1.2.840.113556.1.4.529' = 'Extended DN'
    '1.2.840.113556.1.4.619' = 'Lazy commit'
    '1.2.840.113556.1.4.800' = 'Active Directory >= Windows 2000'
    '1.2.840.113556.1.4.801' = 'SD flags'
    '1.2.840.113556.1.4.805' = 'Tree delete'
    '1.2.840.113556.1.4.906' = 'Microsoft large integer'
    '1.2.840.113556.1.4.1302' = 'Microsoft OID used with DEN Attributes'
    '1.2.840.113556.1.4.1338' = 'Verify name'
    '1.2.840.113556.1.4.1339' = 'Domain scope'
    '1.2.840.113556.1.4.1340' = 'Search options'
    '1.2.840.113556.1.4.1341' = 'RODC DCPROMO'
    '1.2.840.113556.1.4.1413' = 'Permissive Modify'
    '1.2.840.113556.1.4.1670' = 'Active Directory (v5.1)>= Windows 2003'
    '1.2.840.113556.1.4.1781' = 'Microsoft LDAP fast bind extended request'
    '1.2.840.113556.1.4.1791' = 'NTLM Signing and Sealing'
    '1.2.840.113556.1.4.1851' = 'ADAM / AD LDS Supported'
    '1.2.840.113556.1.4.1852' = 'Quota Control'
    '1.2.840.113556.1.4.1880' = 'ADAM Digest'
   # '1.2.840.113556.1.4.1852' = 'Shutdown Notify'
    '1.2.840.113556.1.4.1920' = 'Partial Secrets'
    '1.2.840.113556.1.4.1935' = 'Active Directory (v6.0) >= Windows 2008'
    '1.2.840.113556.1.4.1947' = 'Force Update'
    '1.2.840.113556.1.4.1948' = 'Range Retrieval No Error'
    '1.2.840.113556.1.4.2026' = 'Input DN'
    '1.2.840.113556.1.4.2064' = 'Show Recycled'
    '1.2.840.113556.1.4.2065' = 'Show Deactivated Link'
    '1.2.840.113556.1.4.2080' = 'Active Directory (v6.1) >= Windows 2008 R2'
}

# Forest level diagram reports can be enabled here. You can also just enable the source file
# generation for input into dot.exe or the graphviz gui at another workstation.
$AD_CreateDiagramSourceFiles = $ExportGraphvizDefinitionFiles
$AD_CreateDiagrams = $false
$Graphviz_Path = ''
