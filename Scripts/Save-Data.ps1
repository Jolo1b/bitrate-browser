function Save-Data {
    param ($data, [string] $title)
    Add-Type -AssemblyName System.Windows.Forms
    $SaveFileBrowser = [System.Windows.Forms.SaveFileDialog]

    $res = [System.Windows.MessageBox]::Show("You want to save data", $title, "YesNo", "Question")
    if($res -eq "Yes") {
        $SaveFileDialog = New-Object $SaveFileBrowser
        $SaveFileDialog.FileName = "Data"
        $SaveFileDialog.DefaultExt = ".csv"
        $SaveFileDialog.Filter = "CSV Files (*.csv)|*.csv|XML Files (*.xml)|*.xml"
     
        $res = $SaveFileDialog.ShowDialog()
        if($res -eq "Ok"){
            [string] $filePath = $SaveFileDialog.FileName
            $ext = [System.IO.Path]::GetExtension($filePath)

            if ($ext -eq ".csv"){
                $data | Export-Csv $filePath -NoTypeInformation
            } elseif($ext -eq ".xml"){
                $data | Export-Clixml $filePath -NoTypeInformation
            }
        }
    }
}