function Get-Bitrate {
    param($Path, $Type, $Force)

    # building command
    [string] $command = "Get-ChildItem -Path $Path\* -Recurse -ErrorAction Ignore -Include ("
    $command += if($Type -eq "all") { 
        foreach($item in $global:items) {"'$item',"} 
    } else {"'$Type'"}
    $command = $command.TrimEnd(", ") + ")" 
    $command += if($Force) {"-Force"}

    $AllMusicFiles = Invoke-Expression "$command"
    if($null -eq $AllMusicFiles) { return $null }
    
    # setup shell
    $shell = New-Object -ComObject Shell.Application
    [int] $bitrateAttribute = 28
    $filesPropertyObject = [System.Collections.Generic.List[psobject]]::new()

    foreach($file in $AllMusicFiles){
        $dirObject = $shell.NameSpace($file.Directory.FullName)
        $fileObject = $dirObject.ParseName($file.Name)

        # bitrate acquisition
        $bitrateStr = $dirObject.GetDetailsOf($fileObject, $bitrateAttribute)
        if($bitrateStr -match "\d+") { [int] $bitrate = $Matches[0] }
        else { $bitrate = -1 }

        $filesPropertyObject.Add((New-Object -TypeName psobject -Property @{
            Name = $file.Name;
            Path = $file.FullName;
            Bitrate = $bitrate;
            Size = $fileObject.Size;
            Extension = $file.Extension
        }))

        Remove-Variable -Force -Name dirObject
        Remove-Variable -Force -Name fileObject
        Remove-Variable -Force -Name bitrateStr
        [System.GC]::Collect()
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell)
    return $filesPropertyObject
}