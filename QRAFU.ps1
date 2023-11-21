#######################################

<#
.SYNOPSIS
  Qualys Remote Auto Force Updater.ps1

.DESCRIPTION
  This Script connects to MECM and checks if PC has any pending updates in softwear center. If it does have updates, It will go ahead and force install the updates.  

.PARAMETER File Name
  This Script takes File name as parameter

.OUTPUTS
  Final Output is stored with name "Qualys <Execution Date> v2.xlsx" Under Current Date Folder.

.NOTES
  Version:        1.0
  Author:         Anagh Patel
  Creation Date:  August ‎29, ‎2023
  
.EXAMPLE
  CMDLine Input: QRS.ps1 "Qualys 11.09.2023 v2.xlsx"
#>

#######################################


#takes file name as argument of this script
$file=$args[0]
#saves current path
$cpath = $pwd
#var to save current date in specific formate
$date = Get-Date -Format "MM.dd.yyyy"  

# Self-elevate the script if required
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
  Exit
 }
}

#opening Excel file
$ExcelObj = New-Object -comobject Excel.Application

#Update the Location of the Excel Report 
$ExcelWorkBook = $ExcelObj.Workbooks.Open("$pwd\$date\$file")

#Update the name of the Excel sheet 
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item(1)
$ExcelWorkSheet.Columns.Item(6).Rows.Item(1) = "Update Info"

# Get the number of filled in rows in the XLSX worksheet
$rowcount=$ExcelWorkSheet.UsedRange.Rows.Count

# Loop through all rows in Column 1 starting from Row 2 (these cells contain the domain usernames)
for($i=2;$i -le $rowcount;$i++){
# Get the name of machine from spreadsheet
$MachinesToPatch=$ExcelWorkSheet.Columns.Item(3).Rows.Item($i).Text

if($nameexists -ne $MachinesToPatch){

#checks if device is online or offline
If (Test-Connection -ComputerName $MachinesToPatch -Count 1 -Quiet)
   {
      Write-Host "$MachinesToPatch is Online"
      try{

$MachinesToPatch | % {
[System.Management.ManagementObject[]] $CMMissingUpdates = @(get-wmiobject -query "SELECT * FROM CCM_SoftwareUpdate WHERE ComplianceState = '0'" -namespace "ROOT\ccm\ClientSDK" -ComputerName $_)
write-host "Found $($CMMissingUpdates.count) updates for $_"
$ExcelWorkSheet.Columns.Item(6).Rows.Item($i) = "No Update Found"
$forcedUpdate = 0


if ($CMMissingUpdates.count)
{
    $CMInstallMissingUpdates = (Get-WmiObject -ComputerName $_ -Namespace 'root\ccm\clientsdk' -Class 'CCM_SoftwareUpdatesManager' -List).InstallUpdates($CMMissingUpdates)
    write-host "Forcing $($CMMissingUpdates.count) updates for $_"
    $ExcelWorkSheet.Columns.Item(6).Rows.Item($i) = "Forcing $($CMMissingUpdates.count) updates for $_"
    $forcedUpdate = 1
}
}
}
Catch [System.Exception]{
    Write-Host "Error" -BackgroundColor Red -ForegroundColor Yellow
    $_.Exception.Message
}
   } else {
      Write-Host "$MachinesToPatch is offline"
      # No soup for you...
   }

}
$nameexists = $MachinesToPatch
if($forcedUpdate -ne  1){
$ExcelWorkSheet.Columns.Item(6).Rows.Item($i) = "No Update Found"
}
else{
$ExcelWorkSheet.Columns.Item(6).Rows.Item($i) = "Forcing $($CMMissingUpdates.count) updates for $_"
}


$ExcelWorkBook.Save()
}
# Save the XLS file and close Excel 
$ExcelWorkBook.close($true) 

#copies the final file to new file for next operaion 
$cfName = "Qualys $date v2.xlsx"
$fName = "Qualys $date v3.xlsx" 

Copy-Item "$cpath\$date\$cfName" -Destination "$cpath\$date\$fName" 