Function ConvertTo-HashArray
{
    <#
    .SYNOPSIS
    Convert an array of objects to a hash table based on a single property of the array. 
    
    .DESCRIPTION
    Convert an array of objects to a hash table based on a single property of the array.
    
    .PARAMETER InputObject
    An array of objects to convert to a hash table array.

    .PARAMETER PivotProperty
    The property to use as the key value in the resulting hash.
    
    .PARAMETER LookupValue
    Property in the psobject to be the value that the hash key points to in the returned result. If not specified, all properties in the psobject are used.

    .EXAMPLE
    $DellServerHealth = @(Get-DellServerhealth @_dellhardwaresplat)
    $DellServerHealth = ConvertTo-HashArray $DellServerHealth 'PSComputerName'

    Description
    -----------
    Calls a function which returns a psobject then converts that result to a hash array based on the PSComputerName
    
    .NOTES
    Author:
    Zachary Loeber
    
    Version Info:
    1.1 - 11/17/2013
        - Added LookupValue Parameter to allow for creation of one to one hashs
        - Added more error validation
        - Dolled up the paramerters
        
    .LINK 
    http://www.the-little-things.net 
    #> 
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   HelpMessage='A single or array of PSObjects',
                   Position=0)]
        [AllowEmptyCollection()]
        [PSObject[]]
        $InputObject,
        
        [Parameter(Mandatory=$true,
                   HelpMessage='Property in the psobject to be the future key in a returned hash.',
                   Position=1)]
        [string]$PivotProperty,
        
        [Parameter(HelpMessage='Property in the psobject to be the value that the hash key points to. If not specified, all properties in the psobject are used.',
                   Position=2)]
        [string]$LookupValue = ''
    )

    BEGIN
    {
        #init array to dump all objects into
        $allObjects = @()
        $Results = @{}
    }
    PROCESS
    {
        #if we're taking from pipeline and get more than one object, this will build up an array
        $allObjects += $inputObject
    }

    END
    {
        ForEach ($object in $allObjects)
        {
            if ($object -ne $null)
            {
                try
                {
                    if ($object.PSObject.Properties.Match($PivotProperty).Count) 
                    {
                        if ($LookupValue -eq '')
                        {
                            $Results[$object.$PivotProperty] = $object
                        }
                        else
                        {
                            if ($object.PSObject.Properties.Match($LookupValue).Count)
                            {
                                $Results[$object.$PivotProperty] = $object.$LookupValue
                            }
                            else
                            {
                                Write-Warning -Message ('ConvertTo-HashArray: LookupValue Not Found - {0}' -f $_.Exception.Message)
                            }
                        }
                    }
                    else
                    {
                        Write-Warning -Message ('ConvertTo-HashArray: LookupValue Not Found - {0}' -f $_.Exception.Message)
                    }
                }
                catch
                {
                    Write-Warning -Message ('ConvertTo-HashArray: Something weird happened! - {0}' -f $_.Exception.Message)
                }
            }
        }
        $Results
    }
}