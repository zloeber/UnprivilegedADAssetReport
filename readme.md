#Unprivileged AD Asset Report

With a non-admin account on domain joined computer this script will attempt to gather as much information as possible about the current Active Directory environment.

##Description
With a non-admin account on domain joined computer this script will attempt to gather as much information as possible about the current Active Directory environment.

This project hasn't been updated in over three years and needs some love to be brought up to date and made better. I'm keeping all the old script and project files untouched in its current form under 'Old Project'.

The following information is reported upon:

**FOREST REPORT**

Forest Summary
- Name/Functional Level
- Domain/Site/DC/GC/Exchange/Lync/Pool counts

Forest Features
- Tombstone Lifetime
- Recycle Bin Enabled
- Lync AD Container
- Lync Version
- Exchange Versions

Site Summary
- Site/Subnet/Link/Connection counts
- Sites without site connections count
- Sites without ISTG count
- Sites without subnets count
- Sites wihtout servers count

Exchange Servers
- Organization
- Administrative Group
- Name
- Roles
- Site
- Serial
- Product ID

Lync Elements
- Function (Server/Pool)
- Type (Internal/Edge/Backend/Pool)
- FQDN

Registered DHCP Servers
- Name
- Creation Date

Registered NPS Servers
- Dopmain
- Name
- Type
* Site Information

Site Summary
- Name
- Location
- Domains
- DCs
- Subnets

Site Details
- Name
- Options
- ISTG
- Links
- Bridgeheads
- Adjacencies

Site Subnets
- Subnet
- Site Name
- Location

Site Connections
- Enabled
- Options
- From
- To

Site Links
- Name
- Replication Interval
- Sites
- Change Notification Enabled
* Domain Information

Domains
- Name
- NetBIOS
- Functional Level
- Forest Root
- RIDs Issued
- RIDs Remaining

Domain Password Policies
- Name
- NetBIOS
- Lockout Threshold
- Pass History Length
- Max Pass Age
- Min Pass Age
- Min Pass Length

Domain Controllers
- Domain
- Site
- Name
- OS
- Time
- IP
- GC
- FSMO Roles

Domain Trusts
- Domain
- Trusted Domain
- Direction
- Attributes
- Trust Type
- Created
- Modified

Domain DFS Shares
- Domain
- Name
- DN
- Remote Server

Domain DFSR Shares
- Domain
- Name
- Content
- Remote Servers

Domain Integrated DNS Zones
- Domain
- Partition
- Name
- Record Count
- Created
- Changed

Domain GPOs
- Domain
- Name
- Created
- Changed

Domain Registered Printers
- Domain
- Name
- Server Name
- Share Name
- Location
- Driver Name

Domain Registered SCCM Servers
- Domain
- Name
- Site Code
- Version
- Default MP
- Device MP

Domain Registered SCCM Sites
- Domain
- Name
- Site Code
- Roaming Boundries

**DOMAIN REPORT**

User Account Statistics 1
- Total User Accounts
- Enabled
- Disabled
- Locked
- Password Does Not Expire
- Password Must Change

Account Statistics (count) 2
- Password Not Required
- Dial-in Enabled
- Control Access With NPS
- Unconstrained Delegation
- Not Trusted For Delegation
- No Pre-Auth Required
- Group Statistics

Total Groups
- Built-in
- Universal Security
- Universal Distribution
- Global Security
- Global Distribution
- Domain Local Security
- Domain Local Distribution

Privileged Group Statistics
- Default Priv Group Name
- Current Group Name (if it were changed)
- Member Count

Privileged Group Membership for the following groups:
- Enterprise Admins
- Schema Admins
- Domain Admins
- Administrators
- Cert Publishers
- Account Operators
- Server Operators
- Backup Operators
- Print Operators

Account information for the prior groups:
- Logon ID
- Name
- Password Age (Days)
- Last Logon Date
- Password Does Not Expire
- Password Reversable
- Password Not Required

##Other Information
**Author:** Zachary Loeber

**Website:** http://www.the-little-things.net

**Github:** https://github.com/zloeber/UnprivilegedADAssetReports