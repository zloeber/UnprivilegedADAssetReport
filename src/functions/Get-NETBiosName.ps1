Function Get-NETBiosName ( $dn, $ConfigurationNC ) 
{ 
    try 
    { 
        $Searcher = New-Object System.DirectoryServices.DirectorySearcher  
        $Searcher.SearchScope = "subtree"  
        $Searcher.PropertiesToLoad.Add("nETBIOSName")| Out-Null 
        $Searcher.SearchRoot = "LDAP://cn=Partitions,$ConfigurationNC" 
        $Searcher.Filter = "(nCName=$dn)" 
        $NetBIOSName = ($Searcher.FindOne()).Properties.Item("nETBIOSName") 
        Return $NetBIOSName 
    } 
    catch 
    { 
        Return $null 
    } 
}