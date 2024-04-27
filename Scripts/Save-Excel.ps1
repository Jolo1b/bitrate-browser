function Save-Excel {
    param($data, [string] $path, $styles)
    if($null -eq $styles.Style){ $styles.Style = "None" }

    function Save-File {
        Import-Module ImportExcel
        $data | Export-Excel $filePath -AutoSize -TableStyle $styles.Style
    }

    try {
        Save-File    
    } catch {
        [string] $message = "If you want to save the results directly to xlsx format, " +
        "you need to install the 'ImportExcel' library from the PS Gallery or" + 
        " Github: https://github.com/dfinke/ImportExcel`nIf you don't want to do that," +
        " you can save them in csv (recommended) or xml format, which you can open in Excel."
        [string] $question = "Do you want to install ImportExcel right now from PSGallery?"

        [System.Windows.Forms.MessageBox]::Show($message, $title, "Ok", "Error")
        $choise = [System.Windows.Forms.MessageBox]::Show($question, $title, "YesNo", "Question")

        if($choise -eq "Yes"){
            $arguments = "[System.Console]::Title = 'Installing dependencies' ;" +
            " Install-Module ImportExcel -Force ; "
            " Read-Host 'please close this window'"
            Start-Process powershell -ArgumentList $arguments -Verb RunAs -Wait
            Save-File
        }
    }
}