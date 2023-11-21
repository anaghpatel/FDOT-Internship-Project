#######################################

<#
.SYNOPSIS
  Qualys Report Summerizer.ps1

.DESCRIPTION
  This script summerizes the qualys report. Summary gives Count of Vulnerabilities group by name of the vulnerability.

.PARAMETER File Name
  This Script takes File name as parameter.

.OUTPUTS
  Final Output is stored with name "Qualys <Execution Date> Report Summery.xlsx" Under Current Date Folder.

.NOTES
  Version:        1.0
  Author:         Anagh Patel
  Creation Date:  August ‎16, ‎2023

.EXAMPLE
  CMDLine Input: QRS.ps1 "Qualys 11.09.2023.xlsx"
#>

#######################################

#takes file name as argument of this script
$file=$args[0]
#saves current path
$cpath = $pwd
#var to save current date in specific formate
$date = Get-Date -Format "MM.dd.yyyy" 

#copies the final file to new file for next operaion 
$cfName = "Qualys $date.xlsx"
$fName = "Qualys $date Report Summery.xlsx" 

Copy-Item "$cpath\$date\$cfName" -Destination "$cpath\$date\$fName"

$VulList = @()

function Create-Object ($QiD, $VulName, $VulCount){
#defining properties

$object = New-Object -TypeName PSObject -Property $properties

$object | Add-Member -MemberType NoteProperty -Name QiD -Value $QiD
$object | Add-Member -MemberType NoteProperty -Name VulName -Value $VulName
$object | Add-Member -MemberType NoteProperty -Name VulCount -Value $VulCount

return $object

}



# Open an Excel workbook first:
$ExcelObj = New-Object -comobject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open("$cpath\$date\$fName")
$LastSheet = $ExcelWorkBook.sheets|Select -Last 1
$ExcelWorkSheet2 = $ExcelWorkBook.Sheets.add($LastSheet)
$ExcelWorkSheet2.name ="Summary"

$ExcelWorkSheet1 = $ExcelWorkBook.Sheets.Item(2)
$ExcelWorkSheet2 = $ExcelWorkBook.Sheets.Item(1)

#sorts list by title of vulnerablities.
$objRange       = $ExcelWorkSheet1.UsedRange
$objRange1      = $ExcelWorkSheet1.range("E1")
$ExcelWorkSheet1.Sort.SortFields.Clear()

[void] $ExcelWorkSheet1.Sort.SortFields.Add($objRange1,0,1,0)

$ExcelWorkSheet1.sort.setRange($objRange)  # define the range to sort
$ExcelWorkSheet1.sort.header = 1      # range has a header
$ExcelWorkSheet1.sort.orientation = 1
$ExcelWorkSheet1.sort.apply()


# Get the number of filled in rows in the XLSX worksheet
$rowcount=$ExcelWorkSheet1.UsedRange.Rows.Count
$ExcelWorkSheet2.Columns.Item(1).Rows.Item(1) = "Number of Times"
$ExcelWorkSheet2.Columns.Item(2).Rows.Item(1) = "Qualys Vulnerability ID"
$ExcelWorkSheet2.Columns.Item(3).Rows.Item(1) = "Vulnerability Name"


# Loop through all rows in Column 1 starting from Row 2 (these cells contain the domain usernames)
for($i=2;$i -le $rowcount;$i++){
$Vcount = 0
# Getting QualysID of vulnerablity
$ViD =$ExcelWorkSheet1.Columns.Item(4).Rows.Item($i).Text
if($ViDexists -ne $ViD){

# Getting Name of vulnerablity
$VName =$ExcelWorkSheet1.Columns.Item(5).Rows.Item($i).Text

#creating object that stores vulnerablity info 
$Vulitem = Create-Object -QiD $ViD -VulName $VName -VulCount $Vcount

#adding object to list
$VulList += $Vulitem

}
#if object already exists then it just adds count
$ViDexists = $ViD
$Vulitem.VulCount += 1

}

# Prints the list
#$VulList

$Counter = 1
 ForEach($Vulitem in $VulList) {

$ExcelWorkSheet2.Columns.Item(1).Rows.Item($Counter) = $Vulitem.VulCount
$ExcelWorkSheet2.Columns.Item(2).Rows.Item($Counter) = $Vulitem.QiD
$ExcelWorkSheet2.Columns.Item(3).Rows.Item($Counter) = $Vulitem.VulName
$Counter++
}
#sorts list by title of vulnerablities.
$objRange       = $ExcelWorkSheet2.UsedRange
$objRange1      = $ExcelWorkSheet2.range("A1")
$ExcelWorkSheet2.Sort.SortFields.Clear()

[void] $ExcelWorkSheet2.Sort.SortFields.Add($objRange1,0,2,0)

$ExcelWorkSheet2.sort.setRange($objRange)  # define the range to sort
#$ExcelWorkSheet2.sort.header = 1      # range does not have a header
$ExcelWorkSheet2.sort.orientation = 1
$ExcelWorkSheet2.sort.apply()
$ExcelWorkBook.Save()
# Save the XLS file and close Excel
$ExcelWorkBook.close($true)