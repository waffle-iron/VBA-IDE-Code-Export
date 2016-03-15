$excel = New-Object -ComObject Excel.Application

$excel.Visible = $true
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.WorkSheets.item(1)

# 1 = Module
# 2 = Class
# 3 = From

$xlmodule = $workbook.VBProject.VBComponents.Add(1)
$xlmodule.Name = "menuModule"

$content = [IO.File]::ReadAllText(".\menuModule.bas")

$xlmodule.CodeModule.AddFromString($content)


$xlmodule = $workbook.VBProject.VBComponents.Add(2)
$xlmodule.Name = "clsVBECmdHandler"

$content = [IO.File]::ReadAllText(".\clsVBECmdHandler.cls")

$xlmodule.CodeModule.AddFromString($content)
