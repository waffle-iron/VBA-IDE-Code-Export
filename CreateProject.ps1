# Set the Excel object
$excel = New-Object -ComObject Excel.Application

# Set visability and turn off alert dialogs
$excel.Visible = $true
$excel.DisplayAlerts = $false

# Create a workbook and add a worksheet
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.WorkSheets.item(1)

# 1 = Module
# 2 = Class
# 3 = From

# Add vba module
$xlmodule = $workbook.VBProject.VBComponents.Add(1)
$xlmodule.Name = "menuModule"

# read text from menuModule.bas
$content = [IO.File]::ReadAllText(".\menuModule.bas")

# add it to the newly created module
$xlmodule.CodeModule.AddFromString($content)

# add a class object for clsVBECmdHandler.cls to be read into
$xlmodule = $workbook.VBProject.VBComponents.Add(2)
$xlmodule.Name = "clsVBECmdHandler"

# read text from clsVBECmdHandler
$content = [IO.File]::ReadAllText(".\clsVBECmdHandler.cls")

# add the text to the newly created module
$xlmodule.CodeModule.AddFromString($content)
