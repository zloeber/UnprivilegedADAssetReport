#Version History

**1.8 - 12/23/2016**
- Major rehaul of the code structure but no major feature/functionality in this release

**1.7 - 02/13/2014**
- New save/load functionality! With a switch you can export all collected data
  to xml for later report processing.
- Fixed domain user priveleged report to show lastlogontimestamp as 'never' in html
  report
- Added change notification attribute to site link report section
- Small modification to Format-HTMLTable function to catch errors when processing empty tables
- Slight code clean up
- Fixed issue with domain report count of passwords set to never expire.

**1.6.1 - 01/15/2014**
- Removed superfluous skipdomainreport and skipforestreport paramenters
- Swapped out Colorize-Table with Format-HTMLTable. This means pretty HTML
  reports on older systems where the Linq assemblies are not available.
- Minor fixes.

**1.6 - 01/10/2014**
- Added registered NPS devices
- Added registered DHCP devices
- Added domain registered print devices
- Added SCCM servers and sites
- Added wrapper parameters to entire script with some most used options for directly
  running the script from a powershell prompt.
- Added ability to prompt for input for all major global variables.
- Fixed verbose calling for priv groups and users
- Updated lastlogontimestamp for user export normalization to show never logged in instead
  of a date from the 1600's.
- Added date translation for account expiration in account normalization.
- Updated ad gathering functions to account for inability to connect to domain and silently exit.
- Slight rearrangement of report sections.

**1.5 - 11/26/2013**
- Added the parameter ForceAnonymous along with the code to force anonymous authentication when 
  sending email reports

**1.4 - 11/21/2013**
- Fixed site connections destiniation server output flaw
- Fixed errors occuring when subnets have no sites
- Fixed a number of other errors and bugs related to my prior addition of Get-ADPathName.
- Fixed issues where phantom domains exist in topology

**1.3 - 11/14/2013**
- Fixed DC count issue
- Some formatting changes
- Added detection for newer versions of exchange schemas
- Changed logic for exchange role detection for 2013 to provide accurate results
- Fixed linq issues when running on windows 2012 servers
- Stopped using builtin -split for ldap paths in favor of a custom function called Get-ADPathName
- Added function for resolving msRTCSIP-PrimaryHomeServer to the user's lync pool name in the CSV 
  export of all users
- More changes to the base functions (more error handling and such)

**1.2 - 11/10/2013**
- Added site summary section
- Fixed some code for when no subnets/sites are returned.
- Fixed site options section (I think)
- Changed 'AllowEmptyReport' Section element to saner name of 'ShowSectionEvenWithNoData'
- Commented out write-verbose statements for the report generation portions
- Added timer in the forest data collection routine (as it was taking way too long to process), found
  pulling all properties in the Search-AD function was a real drag so I manually defined all the properties
  to gather where needed. Should speed things up considerably.
- Fixed recycle bin detection
- Prettied up the DC report section to better show FSMO roles and GCs
- Changed the trusts attribute detection to be an enumeration instead
- Mild changes to the base report generation functions.
- Added Exchange Federations section

**1.1 - 11/02/2013**
- Added domain level reporting
- Added AD Integrated Zone information to forest reports
- Added GPO information to forest reports
- Fixed a ton of Powershell V2 related issues

**1.0 - 10/15/2013**
- Initial release of forest level report