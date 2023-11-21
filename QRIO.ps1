#######################################

<#
.SYNOPSIS
  Qualys Report Information Obtainer.ps1

.DESCRIPTION
  This Script gets device information from MECM that helps to identify who does that device belog to and potentially where it is located.  

.PARAMETER File Name
  This Script takes File name as parameter

.OUTPUTS
  Final Output is stored with name "Qualys <Execution Date> v1.xlsx" Under Current Date Folder.

.NOTES
  Version:        1.0
  Author:         Anagh Patel
  Creation Date:  July ‎18, ‎2023
  
.EXAMPLE
  CMDLine Input: QRS.ps1 "Qualys 11.09.2023 v1.xlsx"
#>

#######################################

# This Script runs through all given list of devices and gets their device description from MECM 

Import-Module ConfigurationManager 
#takes file name as argument of this script
$file=$args[0]
#saves current path
$cpath = $pwd
#var to save current date in specific formate
$date = Get-Date -Format "MM.dd.yyyy" 


# Open an Excel workbook first: 

$ExcelObj = New-Object -comobject Excel.Application 

#Update the Location of the Excel Report 

$ExcelWorkBook = $ExcelObj.Workbooks.Open("$cpath\$date\$file") 

#Update the name of the Excel sheet 

$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item(1) 

 

# Get the number of filled in rows in the XLSX worksheet 

$rowcount=$ExcelWorkSheet.UsedRange.Rows.Count 

#runs script that connects psconsole to MECM cmdline 

$scriptPath = "$pwd\ISEConnect_DOT.ps1"

. $scriptPath
  

# Loop through all rows in Column 1 starting from Row 2 (these cells contain the domain usernames) 

for($i=2;$i -le $rowcount;$i++){ 

$MECMComputerName=$ExcelWorkSheet.Columns.Item(3).Rows.Item($i).Text 

if($nameexists -ne $MECMComputerName){ 

# Get the values of description from MECM 

$MECMDeviceDescription = Get-CMDevice -Name $MECMComputerName -Resource | Select-Object description 

Write-Output "$MECMComputerName is done" 

} 

$nameexists = $MECMComputerName 

  

# Fill in the cells with the data received from MECM 

$ExcelWorkSheet.Columns.Item(2).Rows.Item($i) = $MECMDeviceDescription.description 

$ExcelWorkBook.Save() 

} 

# Save the XLS file and close Excel 
$ExcelWorkBook.close($true) 

#copies the final file to new file for next operaion 
$cfName = "Qualys $date v1.xlsx"
$fName = "Qualys $date v2.xlsx" 

Copy-Item "$cpath\$date\$cfName" -Destination "$cpath\$date\$fName" 