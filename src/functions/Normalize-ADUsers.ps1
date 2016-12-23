Function Normalize-ADUsers
{
    [cmdletbinding()]
    param
    (
        [Parameter(HelpMessage='User or users to process.',
                   Mandatory=$true,
                   ValueFromPipeline=$true)]
        [psobject[]]$User,
        
        [Parameter(HelpMessage='AD attributes to process.',
                   Mandatory=$true)]
        [string[]]$Attribs
    )

    BEGIN
    {
        $Users = @()
        $LyncPools = Get-LyncPoolAssociationHash | 
                      ConvertTo-HashArray -PivotProperty 'ServiceName' -LookupValue 'PoolName'
    }
    PROCESS
    {
        if ($User -ne $null)
        {
            $Users += $User
        }
    }
    END
    {
        Foreach ($usr in $Users)
        {
            $UserProps = @{}
            Foreach ($Attrib in $Attribs)
            {
                if ($usr.PSObject.Properties.Match($Attrib).Count) 
                {
                    switch ($Attrib) 
                    {
                        'pwdlastset' {                            
                            $AttribVal = [datetime]::FromFileTime([int64]($usr.$Attrib))
                            $PasswordAge=((get-date) - $AttribVal).days
                            $UserProps.Add(
                                'PasswordAge',
                                $PasswordAge
                            )
                            break
                        }
                        'lastlogontimestamp' {
                            $AttribVal = [datetime]::FromFileTime([int64]($usr.$Attrib))
                            if ($AttribVal -match '12/31/1600')
                            {
                                $LogonAge = 'Never'
                                $AttribVal = 'Never'
                            }
                            else
                            {
                                $LogonAge=((get-date) - $AttribVal).days
                            }
                            $UserProps.Add(
                                'DaysSinceLastLogon',
                                $LogonAge
                            )
                            break
                        }
                        { @('badPasswordTime', 'lastlogon') -contains $_ } {
                            $AttribVal = [datetime]::FromFileTime([int64]($usr.$Attrib))
                            break
                        }
                        'accountExpires'{
                            if (($usr.$Attrib -eq 0) -or ($usr.$Attrib -eq '9223372036854775807') -or ($usr.$Attrib -eq '9223372032559808511'))
                            {
                                $AttribVal = 'Never'
                            }
                            else
                            {
                                $AttribVal = [datetime]::FromFileTime([int64]($usr.$Attrib))
                            }
                            break
                        }
                        'msRTCSIP-PrimaryHomeServer' {
                            if ($usr.$Attrib -ne $null)
                            {
                                $AttribVal = $LyncPools[$usr.$Attrib]
                            }
                            else
                            {
                                $AttribVal = $null
                            }
                            $UserProps.Add(
                                'LyncPool',
                                $AttribVal
                            )
                        }
                        default {
                            $AttribVal = $usr.$Attrib
                            break
                        }
                     }

                    $UserProps.Add(
                            $Attrib,
                            $AttribVal
                    )
                } 
                else 
                { 
                    $UserProps.Add(
                            $Attrib,
                            $null
                    )
                }
            }
            New-Object psobject -Property $UserProps
        }
    }
}