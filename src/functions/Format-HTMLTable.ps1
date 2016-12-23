Function Format-HTMLTable 
{
    <# 
    .SYNOPSIS 
        Format-HTMLTable - Selectively color elements of of an html table based on column value or even/odd rows.
     
    .DESCRIPTION 
        Create an html table and colorize individual cells or rows of an array of objects 
        based on row header and value. Optionally, you can also modify an existing html 
        document or change only the styles of even or odd rows.
     
    .PARAMETER InputObject 
        An array of objects (ie. (Get-process | select Name,Company) 
     
    .PARAMETER  Column 
        The column you want to modify. (Note: If the parameter ColorizeMethod is not set to ByValue the 
        Column parameter is ignored)

    .PARAMETER ScriptBlock
        Used to perform custom cell evaluations such as -gt -lt or anything else you need to check for in a
        table cell element. The scriptblock must return either $true or $false and is, by default, just
        a basic -eq comparisson. You must use the variables as they are used in the following example.
        (Note: If the parameter ColorizeMethod is not set to ByValue the ScriptBlock parameter is ignored)

        [scriptblock]$scriptblock = {[int]$args[0] -gt [int]$args[1]}

        $args[0] will be the cell value in the table
        $args[1] will be the value to compare it to

        Strong typesetting is encouraged for accuracy.

    .PARAMETER  ColumnValue 
        The column value you will modify if ScriptBlock returns a true result. (Note: If the parameter 
        ColorizeMethod is not set to ByValue the ColumnValue parameter is ignored).
     
    .PARAMETER  Attr 
        The attribute to change should ColumnValue be found in the Column specified. 
        - A good example is using "style" 

    .PARAMETER  AttrValue 
        The attribute value to set when the ColumnValue is found in the Column specified 
        - A good example is using "background: red;" 
    
    .PARAMETER DontUseLinq
        Use inline C# Linq calls for html table manipulation by default. This is extremely fast but requires .NET 3.5 or above.
        Use this switch to force using non-Linq method (xml) first.
        
    .PARAMETER Fragment
        Return only the HTML table instead of a full document.
    
    .EXAMPLE 
        This will highlight the process name of Dropbox with a red background. 

        $TableStyle = @'
        <title>Process Report</title> 
            <style>             
            BODY{font-family: Arial; font-size: 8pt;} 
            H1{font-size: 16px;} 
            H2{font-size: 14px;} 
            H3{font-size: 12px;} 
            TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;} 
            TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;} 
            TD{border: 1px solid black; padding: 5px;} 
            </style>
        '@

        $tabletocolorize = Get-Process | Select Name,CPU,Handles | ConvertTo-Html -Head $TableStyle
        $colorizedtable = Format-HTMLTable $tabletocolorize -Column "Name" -ColumnValue "Dropbox" -Attr "style" -AttrValue "background: red;" -HTMLHead $TableStyle
        $colorizedtable = Format-HTMLTable $colorizedtable -Attr "style" -AttrValue "background: grey;" -ColorizeMethod 'ByOddRows' -WholeRow:$true
        $colorizedtable = Format-HTMLTable $colorizedtable -Attr "style" -AttrValue "background: yellow;" -ColorizeMethod 'ByEvenRows' -WholeRow:$true
        $colorizedtable | Out-File "$pwd/testreport.html" 
        ii "$pwd/testreport.html"

    .EXAMPLE 
        Using the same $TableStyle variable above this will create a table of top 5 processes by memory usage,
        color the background of a whole row yellow for any process using over 150Mb and red if over 400Mb.

        $tabletocolorize = $(get-process | select -Property ProcessName,Company,@{Name="Memory";Expression={[math]::truncate($_.WS/ 1Mb)}} | Sort-Object Memory -Descending | Select -First 5 ) 

        [scriptblock]$scriptblock = {[int]$args[0] -gt [int]$args[1]}
        $testreport = Format-HTMLTable $tabletocolorize -Column "Memory" -ColumnValue 150 -Attr "style" -AttrValue "background:yellow;" -ScriptBlock $ScriptBlock -HTMLHead $TableStyle -WholeRow $true
        $testreport = Format-HTMLTable $testreport -Column "Memory" -ColumnValue 400 -Attr "style" -AttrValue "background:red;" -ScriptBlock $ScriptBlock -WholeRow $true
        $testreport | Out-File "$pwd/testreport.html" 
        ii "$pwd/testreport.html"

    .NOTES 
        If you are going to convert something to html with convertto-html in powershell v2 there is 
        a bug where the header will show up as an asterick if you only are converting one object property. 

        This script is a modification of something I found by some rockstar named Jaykul at this site
        http://stackoverflow.com/questions/4559233/technique-for-selectively-formatting-data-in-a-powershell-pipeline-and-output-as

        .Net 3.5 or above is a requirement for using the Linq libraries.

    Version Info:
    1.2 - 01/12/2014
        - Changed bool parameters to switch
        - Added DontUseLinq parameter
        - Changed function name to be less goofy sounding
        - Updated the add-type custom namespace from Huddled to CustomLinq
        - Added help messages to fuction parameters.
        - Added xml method for function to use if the linq assemblies couldn't be loaded (slower but still works)
    1.1 - 11/13/2013
        - Removed the explicit definition of Csharp3 in the add-type definition to allow windows 2012 compatibility.
        - Fixed up parameters to remove assumed values
        - Added try/catch around add-type to detect and prevent errors when processing on systems which do not support
          the linq assemblies.
    .LINK 
        http://www.the-little-things.net 
    #> 
    [CmdletBinding( DefaultParameterSetName = "StringSet")] 
    param ( 
        [Parameter( Position=0,
                    Mandatory=$true, 
                    ValueFromPipeline=$true, 
                    ParameterSetName="ObjectSet",
                    HelpMessage="Array of psobjects to convert to an html table and modify.")]
        [Object[]]
        $InputObject,
        
        [Parameter( Position=0, 
                    Mandatory=$true, 
                    ValueFromPipeline=$true, 
                    ParameterSetName="StringSet",
                    HelpMessage="HTML table to modify.")] 
        [string]
        $InputString='',
        
        [Parameter( HelpMessage="Column name to compare values against when updating the table by value.")]
        [string]
        $Column="Name",
        
        [Parameter( HelpMessage="Value to compare when updating the table by value.")]
        $ColumnValue=0,
        
        [Parameter( HelpMessage="Custom script block for table conditions to search for when updating the table by value.")]
        [scriptblock]
        $ScriptBlock = {[string]$args[0] -eq [string]$args[1]}, 
        
        [Parameter( Mandatory=$true,
                    HelpMessage="Attribute to append to table element.")] 
        [string]
        $Attr,
        
        [Parameter( Mandatory=$true,
                    HelpMessage="Value to assign to attribute.")] 
        [string]
        $AttrValue,
        
        [Parameter( HelpMessage="By default the td element (individual table cell) is modified. This switch causes the attributes for the entire row (tr) to update instead.")] 
        [switch]
        $WholeRow,
        
        [Parameter( HelpMessage="If an array of object is converted to html prior to modification this is the head data which will get prepended to it.")]
        [string]
        $HTMLHead='<title>HTML Table</title>',
        
        [Parameter( HelpMessage="Method for table modification. ByValue uses column name lookups. ByEvenRows/ByOddRows are exactly as they sound.")]
        [ValidateSet('ByValue','ByEvenRows','ByOddRows')]
        [string]
        $ColorizeMethod='ByValue',
        
        [Parameter( HelpMessage="Use inline C# Linq calls for html table manipulation by default. Extremely fast but requires .NET 3.5 or above to work. Use this switch to force using non-Linq method (xml) first.")] 
        [switch]
        $DontUseLinq,
        
        [Parameter( HelpMessage="Return only the html table element.")] 
        [switch]
        $Fragment
        )
    
    BEGIN 
    {
        $LinqAssemblyLoaded = $false
        if (-not $DontUseLinq)
        {
            # A little note on Add-Type, this adds in the assemblies for linq with some custom code. The first time this 
            # is run in your powershell session it is compiled and loaded into your session. If you run it again in the same
            # session and the code was not changed at all, powershell skips the command (otherwise recompiling code each time
            # the function is called in a session would be pretty ineffective so this is by design). If you make any changes
            # to the code, even changing one space or tab, it is detected as new code and will try to reload the same namespace
            # which is not allowed and will cause an error. So if you are debugging this or changing it up, either change the
            # namespace as well or exit and restart your powershell session.
            #
            # And some notes on the actual code. It is my first jump into linq (or C# for that matter) so if it looks not so 
            # elegant or there is a better way to do this I'm all ears. I define four methods which names are self-explanitory:
            # - GetElementByIndex
            # - GetElementByValue
            # - GetOddElements
            # - GetEvenElements
            $LinqCode = @"
            public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByIndex(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, int index)
            {
                return doc.Descendants(element)
                        .Where  (e => e.NodesBeforeSelf().Count() == index)
                        .Select (e => e);
            }
            public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByValue(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, string value)
            {
                return  doc.Descendants(element) 
                        .Where  (e => e.Value == value)
                        .Select (e => e);
            }
            public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetOddElements(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element)
            {
                return doc.Descendants(element)
                        .Where  ((e,i) => i % 2 != 0)
                        .Select (e => e);
            }
            public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetEvenElements(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element)
            {
                return doc.Descendants(element)
                        .Where  ((e,i) => i % 2 == 0)
                        .Select (e => e);
            }
"@
            try
            {
                Add-Type -ErrorAction SilentlyContinue `
                -ReferencedAssemblies System.Xml, System.Xml.Linq `
                -UsingNamespace System.Linq `
                -Name XUtilities `
                -Namespace CustomLinq `
                -MemberDefinition $LinqCode
                
                $LinqAssemblyLoaded = $true
            }
            catch
            {
                $LinqAssemblyLoaded = $false
            }
        }
        $tablepattern = [regex]'(?s)(<table.*?>.*?</table>)'
        $headerpattern = [regex]'(?s)(^.*?)(?=<table)'
        $footerpattern = [regex]'(?s)(?<=</table>)(.*?$)'
        $header = ''
        $footer = ''
    }
    PROCESS 
    { }
    END 
    { 
        if ($psCmdlet.ParameterSetName -eq 'ObjectSet')
        {
            # If we sent an array of objects convert it to html first
            $InputString = ($InputObject | ConvertTo-Html -Head $HTMLHead)
        }

        # Convert our data to x(ht)ml 
        if ($LinqAssemblyLoaded)
        {
            $xml = [System.Xml.Linq.XDocument]::Parse("$InputString")
        }
        else
        {
            # old school xml is kinda dumb so we strip out only the table to work with then 
            # add the header and footer back on later.
            $firsttable = [Regex]::Match([string]$InputString, $tablepattern).Value
            $header = [Regex]::Match([string]$InputString, $headerpattern).Value
            $footer = [Regex]::Match([string]$InputString, $footerpattern).Value
            [xml]$xml = [string]$firsttable
        }
        switch ($ColorizeMethod) {
            "ByEvenRows" {
                if ($LinqAssemblyLoaded)
                {
                    $evenrows = [CustomLinq.XUtilities]::GetEvenElements($xml, "{http://www.w3.org/1999/xhtml}tr")    
                    foreach ($row in $evenrows)
                    {
                        $row.SetAttributeValue($Attr, $AttrValue)
                    }
                }
                else
                {
                    $rows = $xml.GetElementsByTagName('tr')
                    for($i=0;$i -lt $rows.count; $i++)
                    {
                        if (($i % 2) -eq 0 ) {
                           $newattrib=$xml.CreateAttribute($Attr)
                           $newattrib.Value=$AttrValue
                           [void]$rows.Item($i).Attributes.Append($newattrib)
                        }
                    }
                }
            }
            "ByOddRows" {
                if ($LinqAssemblyLoaded)
                {
                    $oddrows = [CustomLinq.XUtilities]::GetOddElements($xml, "{http://www.w3.org/1999/xhtml}tr")    
                    foreach ($row in $oddrows)
                    {
                        $row.SetAttributeValue($Attr, $AttrValue)
                    }
                }
                else
                {
                    $rows = $xml.GetElementsByTagName('tr')
                    for($i=0;$i -lt $rows.count; $i++)
                    {
                        if (($i % 2) -ne 0 ) {
                           $newattrib=$xml.CreateAttribute($Attr)
                           $newattrib.Value=$AttrValue
                           [void]$rows.Item($i).Attributes.Append($newattrib)
                        }
                    }
                }
            }
            "ByValue" {
                if ($LinqAssemblyLoaded)
                {
                    # Find the index of the column you want to format 
                    $ColumnLoc = [CustomLinq.XUtilities]::GetElementByValue($xml, "{http://www.w3.org/1999/xhtml}th",$Column) 
                    $ColumnIndex = $ColumnLoc | Foreach-Object{($_.NodesBeforeSelf() | Measure-Object).Count} 
            
                    # Process each xml element based on the index for the column we are highlighting 
                    switch([CustomLinq.XUtilities]::GetElementByIndex($xml, "{http://www.w3.org/1999/xhtml}td", $ColumnIndex)) 
                    { 
                        {$(Invoke-Command $ScriptBlock -ArgumentList @($_.Value, $ColumnValue))} {
                            if ($WholeRow)
                            {
                                $_.Parent.SetAttributeValue($Attr, $AttrValue)
                            }
                            else
                            {
                                $_.SetAttributeValue($Attr, $AttrValue)
                            }
                        }
                    }
                }
                else
                {
                    $colvalindex = 0
                    $headerindex = 0
                    $xml.GetElementsByTagName('th') | Foreach {
                        if ($_.'#text' -eq $Column) 
                        {
                            $colvalindex=$headerindex
                        }
                        $headerindex++
                    }
                    $rows = $xml.GetElementsByTagName('tr')
                    $cols = $xml.GetElementsByTagName('td')
                    $colvalindexstep = ($cols.count /($rows.count - 1))
                    for($i=0;$i -lt $rows.count; $i++)
                    {
                        $index = ($i * $colvalindexstep) + $colvalindex
                        $colval = $cols.Item($index).'#text'
                        if ($(Invoke-Command $ScriptBlock -ArgumentList @($colval, $ColumnValue))) {
                            $newattrib=$xml.CreateAttribute($Attr)
                            $newattrib.Value=$AttrValue
                            try 
                            {
                                if ($WholeRow)
                                {
                                    [void]$rows.Item($i).Attributes.Append($newattrib)
                                }
                                else
                                {
                                    [void]$cols.Item($index).Attributes.Append($newattrib)
                                }
                            }
                            catch
                            {
                                Write-Warning -Message ('Format-HTMLTable: Something weird happened! - {0}' -f $_.Exception.Message)
                            }
                        }
                    }
                }
            }
        }
        if ($LinqAssemblyLoaded)
        {
            if ($Fragment)
            {
                [string]$htmlresult = $xml.Document.ToString()
                if ([string]$htmlresult -match $tablepattern)
                {
                    [string]$matches[0]
                }
            }
            else
            {
                [string]$xml.Document.ToString()
            }
        }
        else
        {
            if ($Fragment)
            {
                [string]($xml.OuterXml | Out-String)
            }
            else
            {
                [string]$htmlresult = $header + ($xml.OuterXml | Out-String) + $footer
                return $htmlresult
            }
        }
    }
}