Function Append-ADUserAccountControl 
{
    <#
        author: Zachary Loeber
        http://support.microsoft.com/kb/305144
        http://msdn.microsoft.com/en-us/library/cc245514.aspx
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(HelpMessage='User or users to process.', Mandatory=$true, ValueFromPipeline=$true)]
        [psobject[]]$User
    )

    Begin {
        Add-Type -TypeDefinition @" 
        [System.Flags]
        public enum userAccountControlFlags {
            SCRIPT  = 0x0000001,
            ACCOUNTDISABLE  = 0x0000002,
            NOT_USED= 0x0000004,
            HOMEDIR_REQUIRED= 0x0000008,
            LOCKOUT = 0x0000010,
            PASSWD_NOTREQD  = 0x0000020,
            PASSWD_CANT_CHANGE  = 0x0000040,
            ENCRYPTED_TEXT_PASSWORD_ALLOWED  = 0x0000080,
            TEMP_DUPLICATE_ACCOUNT = 0x0000100,
            NORMAL_ACCOUNT  = 0x0000200,
            INTERDOMAIN_TRUST_ACCOUNT               = 0x0000800,
            WORKSTATION_TRUST_ACCOUNT               = 0x0001000,
            SERVER_TRUST_ACCOUNT                    = 0x0002000,
            DONT_EXPIRE_PASSWD                      = 0x0010000,
            MNS_LOGON_ACCOUNT                       = 0x0020000,
            SMARTCARD_REQUIRED                      = 0x0040000,
            TRUSTED_FOR_DELEGATION                  = 0x0080000,
            NOT_DELEGATED   = 0x0100000,
            USE_DES_KEY_ONLY= 0x0200000,
            DONT_REQUIRE_PREAUTH                    = 0x0400000,
            PASSWORD_EXPIRED= 0x0800000,
            TRUSTED_TO_AUTH_FOR_DELEGATION          = 0x1000000
        }
"@
        $Users = @()
        $UACAttribs = @(
            'SCRIPT',
            'ACCOUNTDISABLE',
            'NOT_USED',
            'HOMEDIR_REQUIRED',
            'LOCKOUT',
            'PASSWD_NOTREQD',
            'PASSWD_CANT_CHANGE',
            'ENCRYPTED_TEXT_PASSWORD_ALLOWED',
            'TEMP_DUPLICATE_ACCOUNT',
            'NORMAL_ACCOUNT',
            'INTERDOMAIN_TRUST_ACCOUNT',
            'WORKSTATION_TRUST_ACCOUNT',
            'SERVER_TRUST_ACCOUNT',
            'DONT_EXPIRE_PASSWD',
            'MNS_LOGON_ACCOUNT',
            'SMARTCARD_REQUIRED',
            'TRUSTED_FOR_DELEGATION',
            'NOT_DELEGATED',
            'USE_DES_KEY_ONLY',
            'DONT_REQUIRE_PREAUTH',
            'PASSWORD_EXPIRED',
            'TRUSTED_TO_AUTH_FOR_DELEGATION',
            'PARTIAL_SECRETS_ACCOUNT'
        )
    }
    Process {
        $Users += $User
    }
    End {
        Foreach ($usr in $Users) {
            if ($usr.PSObject.Properties.Match('useraccountcontrol').Count) {
                try {
                    $UAC = [Enum]::Parse('userAccountControlFlags', $usr.useraccountcontrol)
                    $UACAttribs | Foreach {
                        Add-Member -InputObject $usr -MemberType NoteProperty -Name $_ -Value ($UAC -match $_) -Force
                    }
                }
                catch {
                    Write-Warning -Message ('Append-ADUserAccountControl: {0}' -f $_.Exception.Message)
                }
            }
            $usr
        }
    }
}