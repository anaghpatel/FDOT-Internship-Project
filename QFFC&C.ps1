#######################################

<#
.SYNOPSIS
  Qualys File Format Converter & Cleaner.ps1

.DESCRIPTION
  This Script takes Qualys csv report and converts it to xlxs format. Creates Folder named <Execution date>(Ex.11.09.2023). Furthermore, It removes unnecessary items from the sheet and makes it more readable. Final output file is saved under created folder.

.PARAMETER File Name
  This Script takes File name as parameter.

.OUTPUTS
  Final Output is stored with name "Qualys <Execution Date>.xlsx" Under Current Date Folder.

.NOTES
  Version:        1.0
  Author:         Anagh Patel
  Creation Date:  Ocotober ‎31, ‎2023
  
.EXAMPLE
  CMDLine Input: QFFC&C.ps1 "Scheduled-Report-Workstations---Patchable-High-Priority-Vulnerabilities-20231108052004.csv"

#>

#######################################

#takes file name as argument of this script
$file=$args[0]
#saves current path
$cpath = $pwd
#var to save current date in specific formate
$date = Get-Date -Format "MM.dd.yyyy" 
$csvFile = "$pwd\$file" 
mkdir $pwd\$date 
$dName = "Qualys $date.xlsx" 
$xlsxFile = "$pwd\$date\$dName" 
$ExcelObj = New-Object -ComObject Excel.Application 
$ExcelWorkBook = $ExcelObj.Workbooks.Open($csvFile) 
$ExcelWorkBook.SaveAs($xlsxFile, 51) 
$ExcelWorkBook.Close()

$ExcelWorkBook = $ExcelObj.Workbooks.Open("$pwd\$date\$dName")
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item(1)

#REMOVING UNNESSASORY ROWS AND COLUMS
for($i =10; $i -gt 0; $i--) {
$ExcelWorkSheet.Range("A$i").EntireRow.Delete()
}
$ExcelWorkSheet.Range("P:P").EntireColumn.Delete()
$ExcelWorkSheet.Range("O:O").EntireColumn.Delete()
$ExcelWorkSheet.Range("N:N").EntireColumn.Delete()
$ExcelWorkSheet.Range("M:M").EntireColumn.Delete()
$ExcelWorkSheet.Range("L:L").EntireColumn.Delete()
$ExcelWorkSheet.Range("K:K").EntireColumn.Delete()
$ExcelWorkSheet.Range("J:J").EntireColumn.Delete()
$ExcelWorkSheet.Range("G:G").EntireColumn.Delete()
$ExcelWorkSheet.Range("F:F").EntireColumn.Delete()
$ExcelWorkSheet.Range("E:E").EntireColumn.Delete()
$ExcelWorkSheet.Range("B:B").EntireColumn.Delete()

#sorts list by decal of computer
$objRange       = $ExcelWorkSheet.UsedRange
$objRange1      = $ExcelWorkSheet.range("C1")
$ExcelWorkSheet.Sort.SortFields.Clear()

[void] $ExcelWorkSheet.Sort.SortFields.Add($objRange1,0,1,0)

$ExcelWorkSheet.sort.setRange($objRange)  # define the range to sort
$ExcelWorkSheet.sort.header = 1      # range has a header
$ExcelWorkSheet.sort.orientation = 1
$ExcelWorkSheet.sort.apply()

$ExcelWorkBook.Save()

# Save the XLS file and close Excel
$ExcelWorkBook.close($true) 

#copies the final file to new file for next operaion 
$cfName = "Qualys $date.xlsx"
$fName = "Qualys $date v1.xlsx" 

Copy-Item "$cpath\$date\$cfName" -Destination "$cpath\$date\$fName" 

