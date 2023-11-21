#######################################

<#
.SYNOPSIS
  Qualys File Format Converter & Cleaner.ps1

.DESCRIPTION
  <Brief description of script>

.PARAMETER File Name
  This Script takes File name as parameter

.OUTPUTS
  <Outputs if any, otherwise state None - example: Log file stored in C:\Windows\Temp\<name>.log>

.NOTES
  Version:        1.0
  Author:         Anagh Patel
  Creation Date:  November ‎09, ‎2023
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

#######################################


# This Script runs through all given list of devices and runs MECM CONFIG action remotly
Import-Module ConfigurationManager
#takes file name as argument of this script
$file=$args[0]
#saves current path
$cpath = $pwd
#var to save current date in specific formate
$date = Get-Date -Format "MM.dd.yyyy" 
$scriptPath = "$pwd\ISEConnect_DOT.ps1"
. $scriptPath

# Open an Excel workbook first:
$ExcelObj = New-Object -comobject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open("$cpath\$date\$file")
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item(1)

# Get the number of filled in rows in the XLSX worksheet
$rowcount=$ExcelWorkSheet.UsedRange.Rows.Count

# Loop through all rows in Column 1 starting from Row 2 (these cells contain the domain usernames)
for($i=2;$i -le $rowcount;$i++){
$MECMComputerName=$ExcelWorkSheet.Columns.Item(3).Rows.Item($i).Text
if($nameexists -ne $MECMComputerName){

###this code does the actions ##########

#Application Deployment Evaluation:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000121}”
write-host "Application Deployment Evaluation Done for $MECMComputerName"


#Discovery data collection Cycle:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000103}”
write-host "Discovery data collection Cycle Done for $MECMComputerName"

#File Collection Cycle:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000104}”
write-host "File Collection Cycle Done for $MECMComputerName"

#Hardware Inventor Cycle:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000001}”
write-host "Hardware Inventor Cycle Done for $MECMComputerName"

#Machine policy retrieval:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000021}”
write-host "Machine policy retrieval Done for $MECMComputerName"

#Machine policy Evaluation Cycle:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000022}”
write-host "Evaluation Cycle Done for $MECMComputerName"

#Software Inventory Cycle:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000102}”
write-host "Software Inventory Cycle Done for $MECMComputerName"

#Software Metering Usage Report Cycle:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000106}”
write-host "Software Metering Usage Report Cycle Done for $MECMComputerName"

#Software Update Deployment Evaluation Cycle:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000114}”
write-host "Software Update Deployment Evaluation Cycle Done for $MECMComputerName"

#Software Update Scan Cycle:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000113}”
write-host "Software Update Scan Cycle Done for $MECMComputerName"

#User policy evaluation cycle:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000027}”
write-host "User policy evaluation cycle Done for $MECMComputerName"

#Windows installer source list update cycle:
Invoke-WmiMethod -ComputerName $MECMComputerName -Namespace root\ccm -Class sms_client -Name TriggerSchedule “{00000000-0000-0000-0000-000000000107}”
write-host "indows installer source list update cycle Done for $MECMComputerName"





Write-Output $MECMComputerName

############ TODO:run force update remotly ################
}
$nameexists = $MECMComputerName
}
# Save the XLS file and close Excel
$ExcelWorkBook.close($true)
#copies the final file to new file for next operaion 
$fName = "Qualys $date v3.xlsx" 
$fName = "Qualys $date v4.xlsx" 

Copy-Item "$cpath\$date\$cfName" -Destination "$cpath\$date\$fName" 