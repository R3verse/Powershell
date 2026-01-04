# Devloper: Max Jensen
# Date 23-09-14
# Purpose of program: Find any server/computer in AD and display either: Running, Stopped or Filtered services.
Set-ExecutionPolicy RemoteSigned


$ComputerNameOne = Read-Host 'Hvilken SERVER/Computer ønsker du at finde Services på?'

Get-Service -ComputerName $ComputerNameOne | Where-Object {$_.status -eq "running"}
Write-Host -ForegroundColor Cyan "`n`rDine muligheder er følgende:`n"
Write-Host -ForegroundColor Red "1) Ja`n"
Write-Host -ForegroundColor Red "2) Nej`n`r"
#Write-Host -ForegroundColor Red "3) Flere Services`n"
$RestartServices = Read-Host 'Vil du genstarte services?'

if($RestartServices -eq "Ja" -or $RestartServices -eq "1"){
    $ComputerNameTwo = Read-Host 'Hvilken SERVER/Computer ønsker du at genstarte service på?'
    $WhichService = Read-Host 'Hvilket Service Navn vil du genstarte?'
    ############## Killing job first
    $timer = 5
$date = Get-Date #Kill that job...!
if (Get-Process *$WhichService*) 
    { Get-Process *$WhichService* | foreach { 
        if((($date - $_.StartTime).seconds) -gt $timer) {
            $procID = $_.id
            Write-Host -ForegroundColor Magenta "Process $procID is running longer than $timer seconds."
            Write-Host -ForegroundColor Green "Killing process $procID.."
            Stop-Process $procID
        }

 } }
 Start-Sleep 5
 Echo "Killing PID... Please wait!"
    Get-Service -Name $WhichService  -ComputerName $ComputerNameTwo | Stop-service -Force
    Get-Service -ComputerName $ComputerNameTwo | Where-Object {$_.DisplayName -like "$WhichService"} | FT -A Status,Name,DisplayName
    Write-Host -ForegroundColor Green "Genstarter Service... vent venligst!`n`r"
    Get-Service -Name $WhichService  -ComputerName $ComputerNameTwo | Start-service
    Write-Host -ForegroundColor Green "Resultat efter genstart:`n"
    Get-Service -ComputerName $ComputerNameTwo | Where-Object {$_.DisplayName -like "$WhichService"} | FT -A Status,Name,DisplayName
    #Invoke-Command -Computername $ComputerName {Restart-Service $WhichService}
    
}elseif($RestartServices -eq "Nej" -or $RestartServices -eq "2" ){
   Write-Host -ForegroundColor DarkGreen "Continueing..!`n`r"
} 
    


Write-Host -ForegroundColor Cyan "Hvilken type service mode vil du finde i?:`n"
Write-Host -ForegroundColor Red "1) Running"
Write-Host -ForegroundColor Red "2) Stopped"
Write-Host -ForegroundColor Red "3) Filtered Stopped"
Write-Host -ForegroundColor Red "4) Filtered Running"
Write-Host -ForegroundColor Red "5) Filtered`n`r"
$Computer = Read-Host 'Hvad vil du finde af muligheder overstående?'
$ComputerName = Read-Host 'Hvilken SERVER/Computer ønsker du at finde Services på?'


if($Computer -eq "Running" -or $Computer -eq "1"){
    Get-Service -ComputerName $ComputerName | Where-Object {$_.status -eq "running"}
}
elseif($Computer -eq "Stopped" -or $Computer -eq "2"){
    Get-Service -ComputerName $ComputerName | Where-Object {$_.status -eq "stopped"}
}elseif($Computer -eq "Filtered" -or $Computer -eq "3"){
Write-Host -ForegroundColor Red "Du kan bruge Wildcard her i denne søgning, som følgende:`n`r"
Write-Host -ForegroundColor Green "*ASP*`n`r"
    $filtered = Read-Host 'Hvad vil du søge efter?'
    Get-Service -ComputerName $ComputerName | Where-Object {$_.DisplayName -like "$filtered"}
}
elseif($Computer -eq "Filtered Running" -or $Computer -eq "4"){
Write-Host -ForegroundColor Red "Du kan bruge Wildcard her i denne søgning, som følgende:`n`r"
Write-Host -ForegroundColor Green "*ASP*`n`r"
    $filtered = Read-Host 'Hvad vil du søge efter?'
    Get-Service -ComputerName $ComputerName | Where-Object {$_.status -eq "Running"} | Where-Object {$_.DisplayName -like "$filtered"} | FT -A Status,Name,DisplayName
}
elseif($Computer -eq "Filtered Stopped" -or $Computer -eq "5"){
Write-Host -ForegroundColor Red "Du kan bruge Wildcard her i denne søgning, som følgende:`n`r"
Write-Host -ForegroundColor Green "*ASP*`n`r"
    $filtered = Read-Host 'Hvad vil du søge efter?'
    Get-Service -ComputerName $ComputerName | Where-Object {$_.status -eq "stopped"} | Where-Object {$_.DisplayName -like "$filtered"}
} elseif($Computer -eq "Find Alle" -or $Computer -eq "6"){

}
else{
    Write-Host -ForegroundColor Cyan "We're done here... Closing program!"
exit
}