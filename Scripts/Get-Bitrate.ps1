function Get-Bitrate {
    param($pathToDir, $fileType, $force)

    # verify if the specified folder exists
    if(-not (Test-Path $pathToDir -PathType Container)){
        Write-Host "Error: invalid path"
        return 0
    }

    $AllMusicFiles = @()
    if($fileType.GetType().Name -eq "ObjectCollection"){
        $AllFiles = $null
        if($force){
            $AllFiles = Get-ChildItem $pathToDir -Recurse -File -Force
        } else {
            $AllFiles = Get-ChildItem $pathToDir -Recurse -File
        }
        foreach($type in $fileType){
            $AllMusicFiles += $AllFiles | Where-Object { "*$($_.Extension)" -eq $type }
        }
    } else { 
        if($force){
            $AllMusicFiles = Get-ChildItem $pathToDir -Recurse -Filter $fileType -File -Force
        } else {
            $AllMusicFiles = Get-ChildItem $pathToDir -Recurse -Filter $fileType -File
        }
    }

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