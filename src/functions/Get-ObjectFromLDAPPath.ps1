Function Get-ObjectFromLDAPPath
{
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="LDAP path.")]
        [string]
        $LDAPPath,
        
        [Parameter(HelpMessage="Translate the ldap type.")]
        [switch]
        $TranslateNamingAttribute
    )
    $output = @()
    $ldaparr = Get-ADPathName $LDAPPath -split
    $regex = [regex]'(?<LDAPType>^.+)\=(?<LDAPName>.+$)'
    $position = 0
    $ldaparr | %{
        if ($_ -match $regex)
        {
            if ($TranslateNamingAttribute)
            {
                switch ($matches['LDAPType']) 
                {
                      'CN' {$_ldaptype = "Common Name"}
                      'OU' {$_ldaptype = "Organizational Unit"}
                      'DC' {$_ldaptype = "Domain Component"}
                   default {$_ldaptype = $matches['LDAPType']}
                }
            }
            else
            {
                $_ldaptype = $matches['LDAPType']
            }
            $objprop = @{
                LDAPType = $_ldaptype
                LDAPName = $matches['LDAPName']
                Position = $position
            }
            $output += New-Object psobject -Property $objprop
            $position++
        }
    }
    Write-Output -InputObject $output
}