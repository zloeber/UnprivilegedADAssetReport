Function New-ZipFile
{
    #.Synopsis
    #  Expand a zip file, ensuring it's contents go to a single folder ...
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true)]
        $ZipFilePath,

        [Parameter(Position=1, Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [Alias("PSPath","Item")]
        [string[]]
        $InputObject = $Pwd,

        [switch]
        $Append,

        # The compression level (defaults to Optimal):
        #   Optimal - The compression operation should be optimally compressed, even if the operation takes a longer time to complete.
        #   Fastest - The compression operation should complete as quickly as possible, even if the resulting file is not optimally compressed.
        #   NoCompression - No compression should be performed on the file.
        [System.IO.Compression.CompressionLevel]$Compression = "Optimal"
    )
    BEGIN
    {
        # Make sure the folder already exists
        [string]$File = Split-Path $ZipFilePath -Leaf
        [string]$Folder = $(if($Folder = Split-Path $ZipFilePath) { Resolve-Path $Folder } else { $Pwd })
        $ZipFilePath = Join-Path $Folder $File
        # If they don't want to append, make sure the zip file doesn't already exist.
        if(!$Append) 
        {
            if(Test-Path $ZipFilePath) 
            { 
                Remove-Item $ZipFilePath 
            }
        }
        $Archive = [System.IO.Compression.ZipFile]::Open( $ZipFilePath, "Update" )
    }
    PROCESS
    {
        foreach($path in $InputObject) 
        {
            foreach($item in Resolve-Path $path) 
            {
                # Push-Location so we can use Resolve-Path -Relative 
                Push-Location (Split-Path $item)
                # This will get the file, or all the files in the folder (recursively)
                foreach($file in Get-ChildItem $item -Recurse -File -Force | % FullName) 
                {
                    # Calculate the relative file path
                    $relative = (Resolve-Path $file -Relative).TrimStart(".\")
                    # Add the file to the zip
                    $null = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive, $file, $relative, $Compression)
                }
                Pop-Location
            }
        }
    }
    END
    {
        $Archive.Dispose()
        Get-Item $ZipFilePath
    }
}