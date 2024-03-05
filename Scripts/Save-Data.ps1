. "$PSScriptRoot\Set-Style.ps1"

function Save-Data {
    param ($data, [string] $title)
    $SaveFileBrowser = [System.Windows.Forms.SaveFileDialog]

    $res = [System.Windows.Forms.MessageBox]::Show("You want to save data", $title, "YesNo", "Question")
    if($res -eq "Yes") {
        $SaveFileDialog = New-Object $SaveFileBrowser
        $SaveFileDialog.FileName = "bangers"
        $SaveFileDialog.DefaultExt = ".csv"
        $SaveFileDialog.Filter = "CSV Files (*.csv)|*.csv|EXCEL Files (*.xlsx)|*.xlsx|XML Files (*.xml)|*.xml"
     
        $res = $SaveFileDialog.ShowDialog()
        if($res -eq "Ok"){
            [string] $filePath = $SaveFileDialog.FileName
            $ext = [System.IO.Path]::GetExtension($filePath)

            if ($ext -eq ".csv"){
                $data | Export-Csv $filePath -NoTypeInformation
            } elseif($ext -eq ".xlsx"){
                Set-Style $data $filePath
            } elseif($ext -eq ".xml"){
                $data | Export-Clixml $filePath
            }
        }
    }
}