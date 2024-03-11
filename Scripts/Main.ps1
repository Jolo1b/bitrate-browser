Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class DisplaySettings
    {
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();
    }
"@

[DisplaySettings]::SetProcessDPIAware()

$FormObj = [System.Windows.forms.Form]
$LabelObj = [System.Windows.forms.Label]
$ButtonlObj = [System.Windows.forms.Button]
$NumericBoxObj = [System.Windows.forms.NumericUpDown]
$GetDirectoryDialog = [System.Windows.Forms.FolderBrowserDialog]

. "$PSScriptRoot\Get-Bitrate.ps1"
. "$PSScriptRoot\Save-Data.ps1"

[int] $winw = 300
[int] $winh = 110
[int] $winhp = $winh / 5
[int] $space = 4
[string] $title = "Bitrate Browser"

function Set-CenterElement {
    param($size) 
    $newPosition = ([int]($winw / 2 - $size / 2))
    return $newPosition
}

$form = New-Object $FormObj
$form.Text = $title
$form.ClientSize = "$winw,$winh"
$form.BackColor = "#8a8a8a"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.StartPosition = "CenterScreen"
$form.TopMost = $true

$minKbpsBoxSize = @(130, $null)
$minKbpsBoxLocation = @($(Set-CenterElement $minKbpsBoxSize[0]), ($winhp * 3 - $space))
$minKbpsBox = New-Object $NumericBoxObj
$minKbpsBox.Minimum = 0
$minKbpsBox.Maximum = 320
$minKbpsBox.Value = 192
$minKbpsBox.Font = New-Object System.Drawing.Font("Arial", 10)
$minKbpsBox.Size = New-Object System.Drawing.Size($minKbpsBoxSize[0], $minKbpsBoxSize[1])
$minKbpsBox.Location = New-Object System.Drawing.Point($minKbpsBoxLocation[0], $minKbpsBoxLocation[1])
$form.Controls.Add($minKbpsBox)

$unitLabelSize = @(50, 30)
$unitLabelLocation = @(($minKbpsBoxLocation[0] + $minKbpsBoxSize[0]), ($minKbpsBoxLocation[1] + 5))
$unitLabel = New-Object $LabelObj
$unitLabel.Text = "kbps"
$unitLabel.Font = New-Object System.Drawing.Font("Arial", 10)
$unitLabel.Size = New-Object System.Drawing.Size($unitLabelSize[0], $unitLabelSize[1])
$unitLabel.Location = New-Object System.Drawing.Point($unitLabelLocation[0], $unitLabelLocation[1])
$form.Controls.Add($unitLabel)

$folderBrowser = New-Object $GetDirectoryDialog
$folderBrowser.Description = "Select Folder"

function Start-Ation {
    $status = $folderBrowser.ShowDialog()
    if($status -eq "OK"){
        # hide window and show "please wait" message
        $form.Visible = $false
        [string] $jobName = "msgJob"
        Start-Job -Name $jobName -ScriptBlock {
            Add-Type -AssemblyName PresentationFramework
            Wait-Event -Timeout 1
            [System.Windows.MessageBox]::Show('Please Wait (press OK)', 'Bitrate Browser', 'Ok', 'Information')
        }

        $data = Get-Bitrate $folderBrowser.SelectedPath | Where-Object {$_.Bitrate -ge $minKbpsBox.Text}  
        Stop-Job -Name $jobName
        Remove-Job -Name $jobName -Force

        if($null -ne $data){
            $data | Out-GridView -Title $title
            Save-Data $data $title    
        } else {
            [System.Windows.Forms.MessageBox]::Show("mp3 files not found!", $title, "Ok", "Error")
        }

    }
}

$searchButtonSize = @(127, 40)
$searchButtonLocation = @($(Set-CenterElement $searchButtonSize[0]), ($winhp - $space))
$searchButton = New-Object $ButtonlObj
$searchButton.Text = "Select Folder"
$searchButton.Font = New-Object System.Drawing.Font("Arial", 10)
$searchButton.BackColor = "#e0e0e0"
$searchButton.Size = New-Object System.Drawing.Size($searchButtonSize[0], $searchButtonSize[1])
$searchButton.Location = New-Object System.Drawing.Point($searchButtonLocation[0], $searchButtonLocation[1])
$searchButton.add_Click({
    Start-Ation
})

$minKbpsBox.Add_KeyPress({
    if($_.KeyChar -eq 13){
        Start-Ation
    }
})

$form.Controls.Add($searchButton)

$form.ShowDialog()
$form.Dispose()