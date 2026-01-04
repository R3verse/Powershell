####### Check hvis mapper eksistere, ellers Opret mapper 
$hostname = hostname
$folderOne = "C:\tmp"
$folderTwo = "C:\tmp\$hostname_processlist"
if((Test-Path -Path "C:\testfolder" -pathtype Container) -eq $False){
Write-Host -ForegroundColor "Green" "Opretter Mappe: $folderOne ...!"
        New-Item -ItemType directory -Path $folderOne
        }else{
       Write-Host -ForegroundColor Cyan "Du har allerede mappen $folderOne...!"
      
     }
    if((Test-Path -Path "C:\tmp\$hostname_processlist" -pathtype Container) -eq $False){
    New-Item -ItemType directory -Path $folderTwo
        Write-Host -ForegroundColor Green "Opretter Mappe: $folderTwo ...!"
        }else{
        Write-Host -ForegroundColor Cyan "Du har allerede mappen $folderTwo...!"
        
     }

Get-Process | Export-Clixml C:\tmp\$hostname_processlist\processlist.xml
# Uncomment this below if you do not have created the file.
#Compare-Object -ReferenceObject (Import-Clixml C:\tmp\$hostname_processlist\processlist.xml) -DifferenceObject (Get-Process) -Property Name | Export-Csv C:\tmp\$hostname_processlist\newprocesses.csv
exit