clear
Get-Process excel | Stop-Process -Force
$FilePath = "D:\folder\ps1\Workorders.xlsx"
$SheetName = "ADGroup"
$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $true
$objExcel.DisplayAlerts = $true
$WorkBook = $objExcel.Workbooks.Open($FilePath)
$WorkSheet = $WorkBook.sheets.item($SheetName)
$WorksheetRange = $workSheet.UsedRange
$R = $WorksheetRange.Rows.Count
$C = $WorksheetRange.Columns.Count
$WorkSheet.Cells.Item(1,$c+1) = "Status"
for ($i=2;$i-lt$R+1;$i++)
{
$EmployeeID = $WorkSheet.Cells.Item($i,8).value2
$GroupName =  $WorkSheet.Cells.Item($i,7).value2
$Error.Clear()
Get-ADUser -Identity $EmployeeID
if($Error)
{
$WorkSheet.Cells.Item($i,$c+1)= "User not available in AD"
}
else
{
$a=(Get-ADGroupMember -Identity "$GroupName").SamAccountName -contains "$EmployeeID"
if($a)
{
$WorkSheet.Cells.Item($i,$c+1)= "User already available in $GroupName"}
else
{ 
$WorkSheet.Cells.Item($i,$c+1)="User need to added in $GroupName" 
}}}
$WorkBook.save()
$objExcel.quit()