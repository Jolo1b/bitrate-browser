function Get-Bitrate {
    param($pathToDir)
    Add-Type -AssemblyName PresentationFramework

    # verify if the specified folder exists
    if(-not (Test-Path $pathToDir -PathType Container)){
        Write-Host "Error: invalid path"
        return 0
    }

    $AllMp3 = Get-ChildItem $pathToDir -Recurse -Filter *.mp3

    if($null -eq $AllMp3) {
        return $null
    }
    
    $shell = New-Object -ComObject Shell.Application
    [int] $bitrateAttribute = 28
    $filesPropertyObject = @()

    foreach($file in $AllMp3){
        $dirObject = $shell.NameSpace($file.Directory.FullName)
        $fileObject = $dirObject.ParseName($file.Name)

        # bitrate acquisition
        $bitrateStr = $dirObject.GetDetailsOf($fileObject, $bitrateAttribute)
        if($bitrateStr -match "\d+") { [int] $bitrate = $Matches[0] }
        else { $bitrate = -1 }

        $filesPropertyObject += New-Object -TypeName psobject -Property @{
            Name = $file.Name;
            Path = $file.FullName;
            Bitrate = $bitrate;
            Size = $fileObject.Size ;
        }   
    }

    return $filesPropertyObject
    
}