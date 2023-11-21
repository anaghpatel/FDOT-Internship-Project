#######################################

<#
.SYNOPSIS
  Qualys Update Status Checker.ps1

.DESCRIPTION
  This Script gets update status from MECM and saves it to the file. 

.PARAMETER File Name
  This Script takes File name as parameter

.OUTPUTS
  Final Output is stored with name "Qualys <Execution Date> updateStatus.csv" Under Current Date Folder.

.NOTES
  Version:        1.0
  Author:         Anagh Patel
  Creation Date:  August ‎29, ‎2023
  
.EXAMPLE
  CMDLine Input: QRS.ps1 "Qualys 11.09.2023 v3.xlsx"
#>

#######################################

#Define array
$Updates = @()
#takes file name as argument of this script
$file=$args[0]
#saves current path
$cpath = $pwd
#var to save current date in specific formate
$date = Get-Date -Format "MM.dd.yyyy" 
  


#opening Excel file
$ExcelObj = New-Object -comobject Excel.Application

#Update the Location of the Excel Report 
$ExcelWorkBook = $ExcelObj.Workbooks.Open("$cpath\$date\$file") 
#Update the name of the Excel sheet 
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item(1)

# Get the number of filled in rows in the XLSX worksheet
$rowcount=$ExcelWorkSheet.UsedRange.Rows.Count

# Loop through all rows in Column 3 starting from Row 2 (these cells contain the decal(System Name) )
for($i=2;$i -le $rowcount;$i++){
    $UpdateCheckMachine=$ExcelWorkSheet.Columns.Item(3).Rows.Item($i).Text
        if($nameexists -ne $UpdateCheckMachine){
        #checks if device is online or offline
 If (Test-Connection -ComputerName $UpdateCheckMachine -Count 1 -Quiet)
   {
      Write-Host "$UpdateCheckMachine is Online"

Try{
#if the machine is online Checking updates in Software Center
    write-host " Checking $($UpdateCheckMachine)"
    $Updates += Invoke-Command -cn $UpdateCheckMachine {
        $Application =  Get-WmiObject -Namespace "root\ccm\clientsdk" -Class CCM_SoftwareUpdate 
        If(!$Application){
                $Object = New-Object PSObject -Property ([ordered]@{      
                        ArticleId         = " - "
                        Software          = " - "
                        State             = " - "
                })
  
                $Object
        }
        Else{
            Foreach ($App in $Application){
  
                $EvState = Switch ( $App.EvaluationState  ) {
                        '0'  { "None" } 
                        '1'  { "Available" } 
                        '2'  { "Submitted" } 
                        '3'  { "Detecting" } 
                        '4'  { "PreDownload" } 
                        '5'  { "Downloading" } 
                        '6'  { "WaitInstall" } 
                        '7'  { "Installing" } 
                        '8'  { "PendingSoftReboot" } 
                        '9'  { "PendingHardReboot" } 
                        '10' { "WaitReboot" } 
                        '11' { "Verifying" } 
                        '12' { "InstallComplete" } 
                        '13' { "Error" }
                        '14' { "WaitServiceWindow" } 
                        '15' { "WaitUserLogon" } 
                        '16' { "WaitUserLogoff" } 
                        '17' { "WaitJobUserLogon" } 
                        '18' { "WaitUserReconnect" } 
                        '19' { "PendingUserLogoff" } 
                        '20' { "PendingUpdate" } 
                        '21' { "WaitingRetry" } 
                        '22' { "WaitPresModeOff" } 
                        '23' { "WaitForOrchestration" } 
  
  
                        DEFAULT { "Unknown" }
                }
  
                $Object = New-Object PSObject -Property ([ordered]@{      
                        ArticleId         = $App.ArticleID
                        Software          = $App.Name
                        State             = $EvState
                         
                })
  
                $Object
            }
        }
  
    } -ErrorAction Stop | select @{n='Name Of Machine';e={$_.pscomputername}},ArticleID,Software,State
}
Catch [System.Exception]{

    $Updates += $Object | select @{n='Name Of Machine';e={$_.pscomputername}},ArticleId,Software,State
    $Object = New-Object PSObject -Property ([ordered]@{      
                        ArticleId         = $_.Exception.Message
                        Software          = " - "
                        State             = " - "
                })
  
                $Object
                
}}
else {
      Write-Host "$UpdateCheckMachine is offline"
      # No soup for you...
   }
}
$nameexists = $UpdateCheckMachine

}

 
#Export results to CSV
$Updates | Export-Csv "$cpath\$date\Qualys 11.09.2023 updateStatus.csv" -Force -NoTypeInformation

