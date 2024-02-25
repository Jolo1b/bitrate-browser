function Get-Bitrate {
    param($pathToDir)
    
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
    [int] $bitrateAttribute = 0
    $filesPropertyObject = @()

    foreach($file in $AllMp3){
        $dirObject = $shell.NameSpace($file.Directory.FullName)
        $fileObject = $dirObject.ParseName($file.Name)

        for([int] $i = 0; -not $bitrateAttribute; $i++){
            $name = $dirObject.GetDetailsOf($dirObject.Items, $i)
            if($name -eq "Bit rate") { $bitrateAttribute = $i }
        }

        # bitrate acquisition
        $bitrateStr = $dirObject.GetDetailsOf($fileObject, $bitrateAttribute)
        if($bitrateStr -match "\d+") { [int] $bitrate = $Matches[0] }
        else { [int] $bitrate = -1 }

        if($bitrate -ne -1){
            $filesPropertyObject += New-Object -TypeName psobject -Property @{
                Name = $file.Name;
                Path = $file.FullName;
                Bitrate = $bitrate;
                Size = $fileObject.Size ;
            }
        }        
    }

    return $filesPropertyObject
    
}