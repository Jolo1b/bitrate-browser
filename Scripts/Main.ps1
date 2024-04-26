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
$selectBoxObj = [System.Windows.forms.ComboBox]
$GetDirectoryDialog = [System.Windows.Forms.FolderBrowserDialog]
$checkBoxObj = [System.Windows.Forms.CheckBox]

. "$PSScriptRoot\Get-Bitrate.ps1"
. "$PSScriptRoot\Save-Data.ps1"

[int16] $winw = 300
[int16] $winh = 120
[string] $title = "Bitrate Browser"
[int16] $margin_y = $winh - $winh + 20
[int16] $margin_x = 10

$form = New-Object $FormObj
$form.Text = $title
$form.ClientSize = "$winw,$winh"
$form.BackColor = "#8a8a8a"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.StartPosition = "CenterScreen"
$form.TopMost = $true

$minKbpsBox = New-Object $NumericBoxObj
$minKbpsBoxSize = @(130, $null)
$minKbpsBoxLocation =  @($margin_x, ($winh - $minKbpsBox.Height - $margin_y - 5))
$minKbpsBox.Minimum = 0
$minKbpsBox.Maximum = 320
$minKbpsBox.Value = 192
$minKbpsBox.Font = New-Object System.Drawing.Font("Arial", 10)
$minKbpsBox.Size = New-Object System.Drawing.Size($minKbpsBoxSize[0], $minKbpsBoxSize[1])
$minKbpsBox.Location = New-Object System.Drawing.Point($minKbpsBoxLocation[0], $minKbpsBoxLocation[1])
$form.Controls.Add($minKbpsBox)

$unitLabel = New-Object $LabelObj
$unitLabelSize = @(50, 30)
$unitLabelLocation = @(($minKbpsBoxLocation[0] + $minKbpsBoxSize[0]), ($minKbpsBoxLocation[1] + 5))
$unitLabel.Text = "kbps"
$unitLabel.Font = New-Object System.Drawing.Font("Arial", 10)
$unitLabel.Size = New-Object System.Drawing.Size($unitLabelSize[0], $unitLabelSize[1])
$unitLabel.Location = New-Object System.Drawing.Point($unitLabelLocation[0], $unitLabelLocation[1])
$form.Controls.Add($unitLabel)

$forceCheckBox = New-Object $checkBoxObj
$forceCheckBoxSize = @(90, 30)
$forceCheckBoxLocation = @(($unitLabelLocation[0] + $unitLabelSize[0] + 5), ($unitLabelLocation[1] - 3))
$forceCheckBox.Font = New-Object System.Drawing.Font("Arial", 10)
$forceCheckBox.Size = New-Object System.Drawing.Size($forceCheckBoxSize[0], $forceCheckBoxSize[1])
$forceCheckBox.Location = New-Object System.Drawing.Point($forceCheckBoxLocation[0], $forceCheckBoxLocation[1])
$forceCheckBox.Text = "Force"
$form.Controls.Add($forceCheckBox)

$fileTypeBox = New-Object $SelectBoxObj
$fileTypeBoxSize = @(130, $null)
$fileTypeBoxLocation = @($margin_x, $margin_y)
$fileTypeBox.Font = New-Object System.Drawing.Font("Arial", 10)
$fileTypeBox.Size = New-Object System.Drawing.Size($fileTypeBoxSize[0], $fileTypeBoxSize[1])
$fileTypeBox.Location = New-Object System.Drawing.Point($fileTypeBoxLocation[0], $fileTypeBoxLocation[1])
$fileTypeBox.Text = "*.mp3"
$fileTypesRange = @("all",
    "*.mp3", 
    "*.m4a", 
    "*.wav", 
    "*.wma", 
    "*.ogg", 
    "*.flac", 
    "*.aac", 
    "*.aiff")
foreach($fileType in $fileTypesRange) { $fileTypeBox.Items.Add($fileType) }
$fileTypeBox.Add_SelectedValueChanged({
    $minKbpsBox.Maximum = if($fileTypeBox.SelectedItem -eq "*.mp3") {320} else {32768}
})
$form.Controls.Add($fileTypeBox)

$folderBrowser = New-Object $GetDirectoryDialog
$folderBrowser.Description = "Select Folder"
function Start-Ation {
    param($itemBox)
    $status = $folderBrowser.ShowDialog()
    if($status -eq "OK"){
        # hide window and show "please wait" message
        $form.Visible = $false
        [string] $jobName = "msgJob"
        Start-Job -Name $jobName -ScriptBlock {
            Add-Type -AssemblyName PresentationFramework
            Wait-Event -Timeout 1
            [System.Windows.MessageBox]::Show(
                "Please Wait (press OK)", 
                "Bitrate Browser", 
                "Ok", 
                "Information")
        }

        $items = $itemBox.Items
        $selectedItem = $itemBox.SelectedItem
        Write-Host "selected item: $selectedItem"
        $fileType = $selectedItem
        $fileType = if($selectedItem -eq "all") {$items}
        $fileType = if($null -eq $selectedItem) {"*.mp3"}

        $data = Get-Bitrate $folderBrowser.SelectedPath $fileType $forceCheckBox.Checked | Where-Object {$_.Bitrate -ge $minKbpsBox.Text}  
        Stop-Job -Name $jobName
        Remove-Job -Name $jobName -Force

        if($null -ne $data){
            $data | Out-GridView -Title $title
            Save-Data $data $title
        } else {
            [string] $message = if($fileTypeBox.SelectedItem -eq "all")  
                {"files not found!"} else {"$($fileType) files not found!"}  

            [System.Windows.Forms.MessageBox]::Show(
                    $message,
                    $title, 
                    "Ok", 
                    "Error")
        }

    }
}

$searchButton = New-Object $ButtonlObj
$searchButtonSize = @(127, 40)
$searchButtonLocation = @(($winw -  $margin_x - $searchButtonSize[0]), $margin_y)
$searchButton.Text = "Select Folder"
$searchButton.Font = New-Object System.Drawing.Font("Arial", 10)
$searchButton.BackColor = "#e0e0e0"
$searchButton.Size = New-Object System.Drawing.Size($searchButtonSize[0], $searchButtonSize[1])
$searchButton.Location = New-Object System.Drawing.Point($searchButtonLocation[0], $searchButtonLocation[1])
$searchButton.add_Click({ Start-Ation $fileTypeBox })

$minKbpsBox.Add_KeyPress({
    if($_.KeyChar -eq 13) { Start-Ation }
})

$form.Controls.Add($searchButton)

$form.ShowDialog()
$form.Dispose()