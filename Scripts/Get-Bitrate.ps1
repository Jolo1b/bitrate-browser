function Get-Bitrate {
    param($pathToDir)
    
    if(-not (Test-Path $pathToDir -PathType Container)){
        Write-Host "Error: invalid path"
        return 0
    }

    if($null -eq $minBitrate){
        $minBitrate = 0
    }

    $shell = New-Object -ComObject Shell.Application
    $AllMp3 = Get-ChildItem $pathToDir -Recurse -Filter *.mp3

    [int] $bitrateAttribute = 0

    $filesPropertyObject = New-Object -TypeName PSCustomObject -Property @{
        Bitrate = New-Object System.Collections.ArrayList
        Name = New-Object System.Collections.ArrayList
        Path = New-Object System.Collections.ArrayList
        Size = New-Object System.Collections.ArrayList
    }

    foreach($file in $AllMp3){
        $dirObject = $shell.NameSpace($file.Directory.FullName)
        $fileObject = $dirObject.ParseName($file.Name)

        for([int] $i = 0; -not $bitrateAttribute; $i++){
            $name = $dirObject.GetDetailsOf($dirObject.Items, $i)
            if($name -eq "Bit rate") { $bitrateAttribute = $i }
        }

        $bitrateStr = $dirObject.GetDetailsOf($fileObject, $bitrateAttribute)
        if($bitrateStr -match "\d+") { [int] $bitrate = $Matches[0] }
        else { [int] $bitrate = -1 }

        if($bitrate -ne -1){
            $filesPropertyObject.Bitrate += $bitrate
            $filesPropertyObject.Name += $file.Name
            $filesPropertyObject.Path += $file.FullName
            $filesPropertyObject.Size += $fileObject.Size
        }        
    }

    return $filesPropertyObject
    
}