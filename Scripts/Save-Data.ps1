function Save-Data {
    param ($data, [string] $title)
    Add-Type -AssemblyName System.Windows.Forms
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
                try {
                    Import-Module ImportExcel
                    $data | Export-Excel $filePath -AutoSize -TableStyle Dark11
                } catch {
                    [string] $message = "If you want to save the results directly to xlsx format, " +
                    "you need to install the 'ImportExcel' library from the PS Gallery or" + 
                    " Github: https://github.com/dfinke/ImportExcel`nIf you don't want to do that," +
                    " you can save them in csv (recommended) or xml format, which you can open in Excel."
                
                    [System.Windows.Forms.MessageBox]::Show($message, $title, "Ok", "Error")
                }
            } elseif($ext -eq ".xml"){
                $data | Export-Clixml $filePath
            }
        }
    }
}