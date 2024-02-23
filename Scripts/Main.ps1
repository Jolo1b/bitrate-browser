Add-Type -AssemblyName System.Windows.Forms

$FormObj = [System.Windows.forms.Form]
$LabelObj = [System.Windows.forms.Label]
$ButtonlObj = [System.Windows.forms.Button]
$TextBoxObj = [System.Windows.forms.TextBox]
$GetDirectoryDialog = [System.Windows.forms.FolderBrowserDialog]

. "$PSScriptRoot\Get-Bitrate.ps1"

[int] $winw = 250
[int] $winh = 90

$form = New-Object $FormObj
$form.Text = "bitrate getter"
$form.ClientSize = "$winw,$winh"
$form.BackColor = "#d0d0d0"

$minKbpsBoxSize = @(100, $null)
$minKbpsBoxLocation = @([int]($winw / 2 - $minKbpsBoxSize[0] / 2), 50)
$minKbpsBox = New-Object $TextBoxObj
$minKbpsBox.Text = "192"
$minKbpsBox.Size = New-Object System.Drawing.Size($minKbpsBoxSize[0], $minKbpsBoxSize[1])
$minKbpsBox.Location = New-Object System.Drawing.Point($minKbpsBoxLocation[0], $minKbpsBoxLocation[1])
$form.Controls.Add($minKbpsBox)

$folderBrowser = New-Object $GetDirectoryDialog
$folderBrowser.Description = "Select Folder"

$searchButtonSize = @(100 ,30)
$searchButtonLocation = @([int]($winw / 2 - $searchButtonSize[0] / 2) ,10)
$searchButton = New-Object $ButtonlObj
$searchButton.Text = "Select Folder"
$searchButton.Size = New-Object System.Drawing.Size($searchButtonSize[0], $searchButtonSize[1])
$searchButton.Location = New-Object System.Drawing.Point($searchButtonLocation[0], $searchButtonLocation[1])
$searchButton.add_Click({
    $status = $folderBrowser.ShowDialog()
    if($status -eq [System.Windows.Forms.DialogResult]::OK){
        $form.close()
        Get-Bitrate $folderBrowser.SelectedPath | Where-Object {$_.Bitrate -ge $minKbpsBox.Text} | Out-GridView
    }
})
$form.Controls.Add($searchButton)

$unitLabelSize = @(50, 30)
$unitLabelLocation = @([int]($winw / 2 + $minKbpsBoxSize[0] / 2), 50)
$unitLabel = New-Object $LabelObj
$unitLabel.Text = "kbps"
$unitLabel.Size = New-Object System.Drawing.Size($unitLabelSize[0], $unitLabelSize[1])
$unitLabel.Location = New-Object System.Drawing.Point($unitLabelLocation[0], ($unitLabelLocation[1] + 3))
$form.Controls.Add($unitLabel)

$form.ShowDialog()
$form.Dispose()