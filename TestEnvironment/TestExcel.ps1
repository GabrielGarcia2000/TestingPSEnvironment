#***********************************************************************
# PowerShell : TestExcel.ps1                                           *
#   Function : Test writing to MS Excel                                *
#            :                                                         *
#***********************************************************************
#                 M O D I F I C A T I O N S                            *
# -- Date -- ---- Name ---- --------- Description -------------------- *
# 11/06/2009 Gabriel Garcia Created.                                   *
#                                                                      *
#***********************************************************************

###
### Example 
### PS C:\> C:\Scripts\TestExcel.ps1
###

### Variables
$CheckTime = get-Date #-Format "yyyy-MM-dd hh:mm:ss"

### Create a new Excel object using COM
$objExcelApp = New-Object -ComObject Excel.Application
$objExcelApp.visible = $false
$objExcelApp.DisplayAlerts = $false

### Add Workbook
$objWorkbook = $objExcelApp.Workbooks.Add()
$objWorksheet = $objExcelApp.Worksheets.Item(1)
$strOutFileName   = "C:\Scripts\TestEnvironment\TestExcel_$(get-date -f yyyy-MM-dd-HHmmss).xlsx"

########################################
### Create and format column headers ###
########################################
$intRow = 1    # Counter variable for rows
$objWorksheet.Cells.Item($intRow,1)  = "My Column 1"
$objWorksheet.Cells.Item($intRow,2)  = "My Column 2"
for ($col = 1; $col -le 2; $col++)  {
    $objWorksheet.Cells.Item($intRow,$col).Font.Bold = $True
    $objWorksheet.Cells.Item($intRow,$col).Interior.ColorIndex = 48
    $objWorksheet.Cells.Item($intRow,$col).Font.ColorIndex = 34
}
$intRow++

#######################
### Write test data ###
#######################
$objWorksheet.Cells.Item($intRow, 1)  = "Column 1 data"
$objWorksheet.Cells.Item($intRow, 2)  = "Column 2 data"

$objWorksheet.UsedRange.EntireColumn.AutoFit()
$objWorkbook.SaveAs($strOutFileName)

##############################
### Exit Excel Application ###
##############################
$objExcelApp.Quit()
