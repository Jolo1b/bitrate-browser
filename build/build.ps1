Set-Location "$PSScriptRoot\..\Scripts"

$items = Get-ChildItem -Recurse -Filter *.ps1

foreach($item in $items) {
    $content = Get-Content $item.Name 
    $files = Get-ChildItem
    $global:goodline = $false
    
    
    "" | Out-file "$PSScriptRoot\dist\bitrate-browser.ps1"
    
    
    foreach($line in $content){
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
}

Set-Location $PSScriptRoot

ps12exe -noConsole -noOutput -inputFile .\dist\bitrate-browser.ps1 -outputFile ".\dist\bitrate browser_x86.exe" -noError -architecture x86 -noVisualStyles -iconFile ..\Assets\Icon.ico
ps12exe -noConsole -noOutput -inputFile .\dist\bitrate-browser.ps1 -outputFile ".\dist\bitrate browser_x64.exe" -noError -architecture x64 -noVisualStyles -iconFile ..\Assets\Icon.ico