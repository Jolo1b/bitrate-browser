#Add-Type -AssemblyName System.Windows.Forms
. "$PSScriptRoot\Get-Bitrate.ps1"

$PathToFiles = Read-Host "Zadaj cestu k priečinku z mp3"

Get-Bitrate $PathToFiles