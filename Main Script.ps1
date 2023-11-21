#######################################

<#
.SYNOPSIS
  Main Script.ps1

.DESCRIPTION
  This is Main Script that executes multiple scripts to automate force update for PC's on Qualys Report 

.PARAMETER File Name
  This Script takes File name as parameter.

.OUTPUTS
  No Output from Main Script but Multiple output from Different Script execution.

.NOTES
  Version:        1.0
  Author:         Anagh Patel
  Creation Date:  ‎November ‎7, ‎2023
  
.EXAMPLE
  & '.\Main Script.ps1' "Scheduled-Report-Workstations---Patchable-High-Priority-Vulnerabilities-20231108052004.csv"
#>

#######################################




#takes file name as argument of this script
$file=$args[0]
#saves current path
$cpath = $pwd
#var to save current date in specific formate
$date = Get-Date -Format "MM.dd.yyyy" 

#First script(Qualys File Format Converter & Cleaner.ps1) that cleans and orginizes given Qaulys Report
Write-Host " Starting  Qualys File Format Converter & Cleaner.ps1"
. "$cpath\QFFC&C.ps1" $file
#switching back to previous path
cd $cpath
Write-Host " Finished Executing Qualys File Format Converter & Cleaner.ps1"

#Second script(Qualys Report Summerizer.ps1) summerizes the Qualys report and provides count of machine by vulneribility 
Write-Host " Starting  Qualys Report Summerizer.ps1"
. "$cpath\QRS.ps1" "Qualys $date.xlsx"
#switching back to previous path
cd $cpath
Write-Host " Finished Executing Qualys Report Summerizer.ps1"

#Third script(Qualys Report Information Obtainer.ps1) that Obtains Device info from MEMC
Write-Host " Starting  Qualys Report Information Obtainer.ps1"
. "$cpath\QRIO.ps1" "Qualys $date v1.xlsx"
#switching back to previous path
cd $cpath
Write-Host " Finished Executing Qualys Report Information Obtainer.ps1"

#Fourth script(Qualys Remote Auto Force Updater.ps1) Obtains number of update available for that PC and initiates Force install those updates
Write-Host " Starting  Qualys Remote Auto Force Updater.ps1"
. "$cpath\QRAFU.ps1" "Qualys $date v2.xlsx" 
#switching back to previous path
cd $cpath
Write-Host " Finished Executing Qualys Remote Auto Force Updater.ps1"

#Fifth script(Qualys Update Status Checker.ps1) checks status of the force update initiated from last script
Write-Host " Starting  Qualys Update Status Checker.ps1"
. "$cpath\QUSC.ps1" "Qualys $date v3.xlsx"
#switching back to previous path
cd $cpath
Write-Host " Finished Executing Qualys Update Status Checker.ps1"

#Sixth script runs action automatically for configuration Manager in control panel
#'$cpath\QARACM.ps1 "Qualys $date.xlsx" '
#switching back to previous path
#cd $cpath
