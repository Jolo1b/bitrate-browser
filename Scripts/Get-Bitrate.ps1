function Get-Bitrate {
    param($pathToDir, $fileType, $force)

    [string] $command = "Get-ChildItem -Recurse -Filter "
    $AllMusicFiles = @()
    if($fileType.GetType().Name -eq "ObjectCollection"){
        foreach($type in $fileType){ $command += "$type " }
    } else {
        $command += "$fileType "
    }

    $command += if($force) {"-Force"}

    Invoke-Expression "$command"
    Write-Host $AllMusicFiles

    if($null -eq $AllMusicFiles) {
        return $null
    }
    
    $shell = New-Object -ComObject Shell.Application
    [int] $bitrateAttribute = 28
    $filesPropertyObject = @()

    foreach($file in $AllMusicFiles){
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
            Size = $fileObject.Size;
            Extension = $file.Extension
        }   
    }

    return $filesPropertyObject
    
}