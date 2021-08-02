# Integer to letter conversion variable
$alph=@()
65..90|foreach-object{$alph+=[char]$_}
65..90|foreach-object{$alph+=([char]65+[char]$_)}

# Sheet, Header, and filter value arrays
$SVar =@("Migration-User List","Migration-PC List","Migration-Shared Mailboxes","Migration-Resource Mailboxes","Migration-Distribution Lists","Migration-HDrives","Migration-Hexion Data Requests","Migration-Printers & Scanners")
$SVar2 =@("Migration-User List","Migration-PC List","Migration-Shared Mailboxes","Migration-Resource Mailboxes","Migration-Distribution Lists","Migration-HDrives","Migration-Hexion Data Requests","Migration-Printers & Scanners")
$OVar =@("vInfo","vSC_VMK")
$HVar = 14,1,8,4,3,7,2,7
$FVar =@("SOL")
$MVar =@("Not Started","In-Progress","Exception")
[array]::Reverse($SVar)
[array]::Reverse($HVar)
# SCRIPT
$MainPath = “C:\Users\bwagrez\Documents\Scripts\Test\Collaboration and Workstation Migration _PRODUCTION.xlsx”
$MainExcel = New-Object -ComObject excel.application
$MainExcel.visible = $true

$OtherPath = “C:\Users\bwagrez\Documents\Scripts\Test\Bakelite ESXI RVTools all-2021_04_09_ 3_08_22.xlsx"
$OtherExcel = New-Object -ComObject excel.application
$OtherExcel.visible = $false

if(Test-Path -Path C:\Users\bwagrez\Documents\Scripts\Test\TestOutput.xlsx -PathType Leaf){
Remove-Item C:\Users\bwagrez\Documents\Scripts\Test\TestOutput.xlsx
}
$SitePath = “C:\Users\bwagrez\Documents\Scripts\Test\TestOutput.xlsx”
$SiteExcel = New-Object -ComObject excel.application
$SiteExcel.visible = $false

$MainWorkbook = $MainExcel.Workbooks.open($MainPath)
$SiteWorkbook = $SiteExcel.Workbooks.Add()



<# Other Sheets #>
$MainWorksheet = $MainWorkbook.WorkSheets.Item($OVar[0]) # SHEET NAME
$MainWorksheet.activate()
$MainWorksheet.UsedRange.AutoFilter(1,$FVar[0])
$ColumnCount = $MainWorksheet.UsedRange.Columns.Count-1
$LastCell = $alph[$ColumnCount] + $MainWorksheet.UsedRange.Rows.Count
$MainRange = $MainWorkSheet.Range("A1:$LastCell")
$MainRange.Copy() | out-null
$SiteWorksheet = $SiteWorkbook.WorkSheets.Add()
$SiteWorksheet.name = "Migration-ESXi Servers" 
$SiteWorksheet.activate()
$SiteWorksheet.Range("A1").PasteSpecial(-4122, $false) | out-null
$LastCell = $alph[$ColumnCount] + $SiteWorksheet.UsedRange.Rows.Count
$ListObject = $SiteWorksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $SiteWorksheet.Range("A1:$LastCell"), $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
$ListObject.TableStyle = "TableStyleMedium1"
$SiteWorksheet.UsedRange.Columns.Autofit() | out-null

$MainWorksheet = $MainWorkbook.WorkSheets.Item($OVar[1]) # SHEET NAME
$MainWorksheet.activate()
$MainWorksheet.UsedRange.AutoFilter(2,"SOL*")
$ColumnCount = $MainWorksheet.UsedRange.Columns.Count-1
$LastCell = $alph[$ColumnCount] + $MainWorksheet.UsedRange.Rows.Count
$MainRange = $MainWorkSheet.Range("A1:$LastCell")
$MainRange.Copy() | out-null
$SiteWorksheet = $SiteWorkbook.WorkSheets.Add()
$SiteWorksheet.name = "Migration-VM Servers" 
$SiteWorksheet.activate()
$SiteWorksheet.Range("A1").PasteSpecial(-4122, $false) | out-null
$LastCell = $alph[$ColumnCount] + $SiteWorksheet.UsedRange.Rows.Count
$ListObject = $SiteWorksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $SiteWorksheet.Range("A1:$LastCell"), $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
$ListObject.TableStyle = "TableStyleMedium1"
$SiteWorksheet.UsedRange.Columns.Autofit() | out-null

<# ------                   ~#>
<# Filtered Sheets #>
for($i=0;$i -lt $SVar.Count;$i++){
$MainWorksheet = $MainWorkbook.WorkSheets.Item($SVar[$i]) # SHEET NAME
$MainWorksheet.activate()
$MainWorksheet.UsedRange.AutoFilter($HVar[$i],$FVar[0]) | out-null #### (COLUMN NUMBER, COLUMN FILTER)
if($MainWorksheet.Name -eq $SVar2[0]){$MainWorksheet.UsedRange.AutoFilter(3,$MVar,$xlFilterValues)}
if($MainWorksheet.Name -eq $SVar2[2]){$MainWorksheet.UsedRange.AutoFilter(9,$MVar,$xlFilterValues)}
if($MainWorksheet.Name -eq $SVar2[3]){$MainWorksheet.UsedRange.AutoFilter(6,$MVar,$xlFilterValues)}
$ColumnCount = $MainWorksheet.UsedRange.Columns.Count-1
$LastCell = $alph[$ColumnCount] + $MainWorksheet.UsedRange.Rows.Count
$MainRange = $MainWorkSheet.Range("A1:$LastCell")
#Write-Output $LastCell
$MainRange.Copy() | out-null
$SiteWorksheet = $SiteWorkbook.WorkSheets.Add()
$SiteWorksheet.name = $MainWorksheet.name 
$SiteWorksheet.activate()
$SiteWorksheet.Range("A1").PasteSpecial(-4122, $false) | out-null
$LastCell = $alph[$ColumnCount] + $SiteWorksheet.UsedRange.Rows.Count
$ListObject = $SiteWorksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $SiteWorksheet.Range("A1:$LastCell"), $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
$ListObject.TableStyle = "TableStyleMedium1"
$SiteWorksheet.UsedRange.Columns.Autofit() | out-null
}
<# ------                   ~#>






<#for ($i=$SVar.Count; $i -ge 1; $i--) {
    $MainWorksheet = $MainWorkbook.WorkSheets.Item($SVar[$i]) # SHEET NAME
    $MainWorksheet.activate()
    $ColumnCount = $MainWorksheet.UsedRange.Columns.Count-1
    $LastCell = $alph[$ColumnCount] + $MainWorksheet.UsedRange.Rows.Count
    $MainWorksheetRange.AutoFilter($Header[$i],$HeaderValues[1])
    $MainRange = $MainWorkSheet.Range("A1:$LastCell")
    $MainRange.Copy() | out-null
    Write-Output $LastCell $i $Header[$i]
    $SiteWorksheet = $SiteWorkbook.WorkSheets.Add()
    $SiteWorksheet.name = $MainWorksheet.name
    $SiteWorksheet.activate()
    $SiteRange = $SiteWorksheet.Range("A1")
    $SiteWorksheet.PasteSpecial($SiteRange)
}#>
$SiteWorkbook.Worksheets.item("Sheet1").Delete()
$SiteWorkbook.SaveAs($SitePath) 
$MainExcel.Quit()
$OtherExcel.Quit()
$SiteExcel.Quit()
spps -n Excel
Remove-Variable -Name SiteExcel
Remove-Variable -Name MainExcel
[gc]::collect()
[gc]::WaitForPendingFinalizers()