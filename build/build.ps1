Set-Location "$PSScriptRoot\..\Scripts"

$mainFile = Get-Content Main.ps1 
$files = Get-ChildItem
$global:goodline = $false


"" | Out-file "$PSScriptRoot\dist\bitrate-browser.ps1"


foreach($line in $mainFile){
    foreach($file in $files){
        if($line -like ". `"`$PSScriptRoot\$($file.Name)`""){
            Write-Host $line
            Get-Content $file.FullName | Out-file "$PSScriptRoot\dist\bitrate-browser.ps1" -Append
            $global:goodline = $true        
        }
    }

    if(-not $global:goodline) {
        $line | Out-file "$PSScriptRoot\dist\bitrate-browser.ps1" -Append 
    } 
    else {
        $global:goodline = $false
    }
}

Set-Location $PSScriptRoot

ps2exe -noConsole -noOutput -inputFile .\dist\bitrate-browser.ps1 -outputFile ".\dist\bitrate browser_x86.exe" -noError -x86 -iconFile ..\Assets\Icon.ico
ps2exe -noConsole -noOutput -inputFile .\dist\bitrate-browser.ps1 -outputFile ".\dist\bitrate browser_x64.exe" -noError -x64 -iconFile ..\Assets\Icon.ico