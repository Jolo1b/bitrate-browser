function Save-Excel {
    param($data, [string] $path, $styles)
    if($null -eq $styles.Style){ $styles.Style = "None" }

    try {
        Import-Module ImportExcel
        $data | Export-Excel $filePath -AutoSize -TableStyle $styles.Style
    } catch {
        [string] $message = "If you want to save the results directly to xlsx format, " +
        "you need to install the 'ImportExcel' library from the PS Gallery or" + 
        " Github: https://github.com/dfinke/ImportExcel`nIf you don't want to do that," +
        " you can save them in csv (recommended) or xml format, which you can open in Excel."
    
        [System.Windows.Forms.MessageBox]::Show($message, $title, "Ok", "Error")
    }
}