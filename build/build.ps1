Set-Location "$PSScriptRoot\..\Scripts"
$items = Get-ChildItem -Recurse -Filter *.ps1

"" | Out-file "$PSScriptRoot\dist\bitrate-browser.ps1"
$mainFileContent = Get-Content "$PSScriptRoot\..\Scripts\Main.ps1"

$red = [char]0x001b + "[31m"
$noColor = [char]0x001b + "[0m"

Write-Host $red"Building..."$noColor

$i = 0
while($i -le 3){
    foreach($item in $items) {
        Write-Host "-> $([char]0x001b)[32m$($item.FullName)"
        $content = Get-Content $item.FullName -Raw

        $mainFileContent = $mainFileContent.Replace(". `"`$PSScriptRoot\$($item.Name)`"", $content)
    }
    $i++
}

$mainFileContent | Set-Content "$PSScriptRoot\dist\bitrate-browser.ps1"
Set-Location $PSScriptRoot

Write-Host $red"Compiling..."$noColor

ps12exe -noConsole -noOutput -inputFile .\dist\bitrate-browser.ps1 -outputFile ".\dist\bitrate browser_x86.exe" -noError -architecture x86 -noVisualStyles -iconFile ..\Assets\Icon.ico
ps12exe -noConsole -noOutput -inputFile .\dist\bitrate-browser.ps1 -outputFile ".\dist\bitrate browser_x64.exe" -noError -architecture x64 -noVisualStyles -iconFile ..\Assets\Icon.ico
ps12exe -noConsole -noOutput -inputFile .\dist\bitrate-browser.ps1 -outputFile ".\dist\bitrate browser_anycpu.exe" -noError -architecture anycpu -noVisualStyles -iconFile ..\Assets\Icon.ico

Write-Host $red"Press any key to continue..."$noColor -NoNewline
[System.Console]::ReadKey()