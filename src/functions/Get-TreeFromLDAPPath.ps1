Function Get-TreeFromLDAPPath
{
    # $Output = [System.Web.HttpUtility]::HtmlDecode(($a | ConvertTo-Html))
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="LDAP path.")]
        [string]
        $LDAPPath,
        
        [Parameter(HelpMessage="Determines the depth a tree node is indented")]
        [int]
        $IndentDepth=1,
        
        [Parameter(HelpMessage="Optional character to use for each newly indented node.")]
        [char]
        $IndentChar = 3,
        
        [Parameter(HelpMessage="Don't remove the ldap node type (ie. DC=)")]
        [Switch]
        $KeepNodeType
     )
    $regex = [regex]'(?<LDAPType>^.+)\=(?<LDAPName>.+$)'
    $ldaparr = Get-ADPathName $LDAPPath -split
    $ADPartCount = $ldaparr.count
    $spacer = ''
    $output = ''
    for ($index = ($ADPartCount); $index -gt 0; $index--) 
    {
        $node = $ldaparr[($index-1)]
        if (-not $KeepNodeType)
        {
            if ($node -match $regex)
            {
                $node = $matches['LDAPName']
            }
        }
        if ($index -eq ($ADPartCount))
        {
            $line = ''
        }
        else
        {
            $line = $IndentChar
            $spacer = $spacer + (' ' * $IndentDepth)
            # This fixes an offset issue
            if ($index -lt ($ADPartCount - 1))
            {
                $spacer = $spacer + ' '
            }
        }
        $line = $spacer + $line + $node + "`n"
        $output = $Output+$line
    }
    [string]$output
}