. "$PSScriptRoot\Save-Excel.ps1"

function Set-Style {
    param($data, [string] $path)
    
    $allStyles = @("None", "Custom")
    for([byte] $i = 1;$i -le 21;$i++){ $allStyles += "Light$i" }
    for([byte] $i = 1;$i -le 28;$i++){ $allStyles += "Medium$i" }
    for([byte] $i = 1;$i -le 11;$i++){ $allStyles += "Dark$i" }
    [int] $winw = 300
    [int] $winh = 80

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Table styles"
    $form.BackColor = "#b9b9b9"
    $form.ClientSize = "$winw,$winh"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.AutoScaleMode = "None"
    $form.StartPosition = "centerScreen"
    $form.TopMost = $true

    $selectStyleWidth = 134
    $selectStyle = New-Object System.Windows.Forms.ComboBox
    $selectStyle.Text = "None"
    $selectStyle.Size = New-Object System.Drawing.Size($selectStyleWidth, 0)
    $selectStyle.Font = New-Object System.Drawing.Font("Arial", 10)
    $selectStyle.Location = New-Object System.Drawing.Point(($winw - $selectStyleWidth - 25), 25)
    foreach($style in $allStyles){
        $selectStyle.Items.Add($style)
    }
    $form.Controls.Add($selectStyle)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "ok"
    $okButton.BackColor = "#eeeeee"
    $okButton.Location = New-Object System.Drawing.Point(30, 24)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Font = New-Object System.Drawing.Font("Arial", 10)
    $okButton.Add_Click({   
        $selectedSyles = New-Object -TypeName psobject -Property @{
            Style = $selectStyle.SelectedItem
        }
        $form.Visible = $false
        $form.Close()       
        Save-Excel $data $path $selectedSyles
    })
    $form.Controls.Add($okButton)

    $form.ShowDialog()
    $form.Dispose()
}