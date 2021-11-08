$SourcePath = "C:\test\"  #This indicated the folder that the .xls files are in
$DestinationPath = "C:\test\upload\"  #This indicated where the converted .txt files will be saved

$files = Get-ChildItem $SourcePath*.xlsx -recurse

$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $false
$Excel.DisplayAlerts = $false

ForEach ($file in $files) {
     Write-Host "Loading File '$($file.Name)'..."
     $WorkBook = $Excel.Workbooks.Open($file.Fullname)
     $NewName = $($file.Name.TrimEnd('.xlsx'))
     $NewFilePath = $DestinationPath + $NewName + ".txt"
     $WorkBook.SaveAs($NewFilepath, 42)   # xlUnicodeText
}

# cleanup
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()