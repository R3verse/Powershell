Write-Host "`r`n"
Write-host -ForegroundColor Cyan "---------------------------------------------`r`n"
Import-Module Activedirectory
Set-ExecutionPolicy RemoteSigned
################# Use these commands or incoporate them!
$path = "\\$($computer.name)\C$\testfolder\SoftwareVersions" + ".csv"


$folderOne = "\\$($computer.name)\C$\testfolder"
$folderTwo = "\\$($computer.name)\C$\testfolder\Testme"
if((Test-Path -Path "C:\testfolder" -pathtype Container) -eq $False){
Write-Host -ForegroundColor "Green" "Opretter Mappe: $folderOne ...!"
        New-Item -ItemType directory -Path $folderOne
        }else{
       Write-Host -ForegroundColor Cyan "Du har allerede mappen $folderOne...!"
      
     }
    if((Test-Path -Path "\\$($computer.name)\C$\testfolder\Testme" -pathtype Container) -eq $False){
    New-Item -ItemType directory -Path $folderTwo
        Write-Host -ForegroundColor Green "Opretter Mappe: $folderTwo ...!"
        }else{
        Write-Host -ForegroundColor Cyan "Du har allerede mappen $folderTwo...!"
        
     }
    
    # $computers = Get-ADComputer jetpc86 -Properties Name,OperatingSystem,ipv4address | Select-Object Name,ipv4address
#$computers = Get-ADComputer -Filter * -Properties Name,OperatingSystem,ipv4address | Select-Object Name,ipv4address 

####### Use this command
Write-Host "`r`n"
Write-host -ForegroundColor Cyan "---------------------------------------------`r`n"
Write-Host -ForegroundColor Red "Husk at indskrive dette i konsollen hvis du vil formate outputtet til CSV!`r`n"
Write-Host "`r`n"
Write-Host -ForegroundColor Yellow 'Husk at sætte variable: $computers = Get-ADComputer -Filter * -Properties Name,OperatingSystem,ipv4address | Select-Object Name,ipv4address'
Write-Host -ForegroundColor Yellow 'For alle computere/servere: Get-Java7Version -Computers $computers -Verbose | select Computer,IPv4Address,Version,File |export-csv  $path -Delimiter ";" -NoType -Encoding UTF8'
Write-Host -ForegroundColor Yellow "For enkelt: Get-ADComputer <REPLACE> -Properties Name,OperatingSystem,ipv4address | Select-Object Name,ipv4address`n"
Write-Host -ForegroundColor Yellow 'For GUI Grid View: Get-Java7Version -Computers $computers -Verbose | select Computer,IPv4Address,Version,File| Out-GridView'
Write-host -ForegroundColor Cyan "---------------------------------------------`r`n"
#############################



#Getting all installed software: Get-WmiObject -Class Win32_Product -ComputerName Jetpc90 |Select IdentifyingNumber,Name,Version| FT -AutoSize
#########################



################


Workflow Get-Java7Version {

[cmdletbinding()]

param(

[psobject[]]$Computers

)

foreach -parallel ($computer in $computers) {

Write-Verbose -Message "Running against $($computer.Name)"

# Check each computer for the info.

$versions = Get-WmiObject -Class CIM_Datafile -Filter "Name=(\\$($computer.Name)\\C$\\Program Files\\Java\\jre7\\bin\\java.dll)" -PSComputerName $($computer.Name).ipv4address -ErrorAction SilentlyContinue | Select-Object -Property name,version
$versionTwo = Get-WmiObject -Class CIM_Datafile -Filter "Name=(\\$($computer.Name)\\C$\\Program Files (x86)\\Java\\jre7\\bin\\java.dll)" -PSComputerName $($computer.Name).ipv4address -ErrorAction SilentlyContinue | Select-Object -Property name,version

if ($versions -or $versionTwo){

# Process each version found

Write-Verbose -Message "Java found on $($computer.Name)"

foreach($version in $versions){

# Create a Custom PSObject with the info to process

$found = InlineScript {

New-Object –TypeName PSObject –Prop @{'Computer'=$Using:computer.Name;

'IPv4Address'=$Using:computer.ipv4address;

'File'=$Using:version.name -or $Using:versionTwo.name;

'Version'=$Using:version.version -or $Using:versionTwo.version 

}

}

# Return the custom object

$found

}

}

else {

Write-Verbose -Message "Java not found in $($computer.Name)"

}

}

} 

################ Start uninstall and new installation!

#First we start a remoting session to the PC
Import-Module ActiveDirectory
Set-ExecutionPolicy RemoteSigned

#$computername = Read-Host "Enter remote computer name"
Write-Host "`r`n"
 write-host -ForegroundColor Cyan "----------------------------------`r`n"
write-host -ForegroundColor Yellow "1) Afinstaller/Installer på en Specifik Computer"
write-host -ForegroundColor Yellow "2) Afinstaller/Installer på Alle mine Servere"
write-host -ForegroundColor Yellow "3) Afinstaller/Installer på Alle mine Computere & Bærbare i SBSComputers"
write-host -ForegroundColor Yellow "4) Afinstaller/Installer på Alle computere/servere i mit AD"
Write-Host "`r`n"
 write-host -ForegroundColor Cyan "----------------------------------`r`n"
 Write-Host -ForegroundColor Red "`n ORB: HUSK AT BRUGE #5 FØRST - DU SKAL BRUGE DET SENERE i BRUGER INPUT!`n"
 write-host -ForegroundColor Cyan "----------------------------------`r`n"
Write-Host -ForegroundColor Green "5) Vis alle java versioner på specifik computer"
Write-host -ForegroundColor Green "6) Afinstaller en gammel java version"
Write-Host -ForegroundColor Green "7) Installer den nyeste Java Version"
write-host -ForegroundColor Yellow "8) Exit program"
Write-Host "`r`n"
$Option = Read-Host "Hvilke muligheder vil du benytte af overstående?"
if($Option -eq "1") {
$UseThis = Read-Host "Hvilken Computer/Server vil du afinstallere/installere på?"
$computernames = Get-ADComputer $UseThis | Select-Object Name
}if($Option -eq "2") {
$computernames = Get-ADComputer -Filter {Where-Object -Match "*server*"} | Select-Object Name
}if($Option -eq "3") {
$computernames = Get-ADComputer -Filter * -SearchBase "OU=SBSComputers,OU=Computers,OU=domain,DC=domain,DC=local" | Select-Object Name
}
if($Option -eq "4") {
$computernames = Get-ADComputer -Filter * | Select-Object Name
}if($Option -eq "5") {
$UseThis = Read-Host "Hvilken Computer/Server vil du Vise alle java versioner på?"
gwmi win32_product -ComputerName $UseThis -filter "Name LIKE '%Java%'" | FT -AutoSize
Write-Host -ForegroundColor Red "`n Hvis du får intet resultat tilbage, men blank skærm, så er det fordi der ikke er Java Installeret!`n"
Write-Host -ForegroundColor Red "`n`r Exiting program...No more stuff to do!"
exit
}if($Option -eq "6") {
$UseThis = Read-Host "Hvilken Computer/Server vil du afinstallere en java version på?"
Write-Host -ForegroundColor Yellow "Eksempel på Identifying number kan findes ved brug af #5 når du starter scriptet!...`n"
$UseIdNum = Read-Host "Hvilket IdentifyingNumber vil du bruge (SKAL MATCHE JAVA VERSIONEN!)"
$UseName = Read-Host "Hvilket Name vil du bruge fra #5?"
$UseVersion = Read-Host "Hvilken Version vil afinstallere - som du kan finde ved brug fra #5?"
$classKey="IdentifyingNumber=`"`{$UseIdNum`}`",Name=`"$UseName`",version=`"$UseVersion`""
([wmi]"\\$UseThis\root\cimv2:Win32_Product.$classKey").Uninstall()
start-sleep 7
Write-Host -ForegroundColor Red "Java $UseVersion er nu afinstalleret på $UseThis. Lukker programmet!..."
exit
}if($Option -eq "7") {
$UseThis = Read-Host "Hvilken Computer/Server vil du Installere den nyeste Java Version på?"
\\COMPUTERNAME\C$\PSTools\PsExec.exe \\$UseThis -u domain\administrator -p K4lEoDu7 msiexec /i "\\SERVER.domain.local\IXP$\SW\JAV00009\IMG\jre1.7.0_67.msi" /qn /norestart /log "\\FILESERVER01\IT\Dokumentation\JavaVersions\Setup_$UseThis.log"
start-sleep 7
Write-Host -ForegroundColor Red "Java $UseVersion er nu Installeret på $UseThis. Lukker programmet!..."
#### Notify user of new installation
$block = { msg console /time:15 "Kære Kollega,`r`r`nVi har installeret Den nyeste version af Java`nSamtidig har vi afinstalleret den gamle!`r`n`rMed venlig hilsen,`rIT Afdelingen"}
Invoke-Command -ComputerName $UseThis -ScriptBlock $block
Write-Host "`r`n"
$Logfolder = " \\FILESERVER01\IT\Dokumentation\JavaVersions\"
Write-Host -ForegroundColor Yellow "Du kan finde alle logs på: $Logfolder`n"
exit
}
if($Option -eq "8") {
Write-Host -ForegroundColor Red "Afslutter program...!"
exit
}


foreach($computer in $computernames) {
$session = New-PSSession -ComputerName $($computer.name)
Enter-PSSession $session 

$GivemeName = $computer.Name

# If there is a previous version installed we need to stop the process
 
if (Get-Process -Name Java -ea SilentlyContinue) {Stop-Process -Name Java -Force}
 
if (Get-Process -Name JP2Launcher -ea SilentlyContinue) {Stop-Process -Name JP2Launcher -Force}
 
if (Get-Process -Name javaw -ea SilentlyContinue) {Stop-Process -Name javaw -Force}
 
#Then we find all objects named Java and uninstall them

$app = Get-WmiObject -ComputerName $GivemeName -Class Win32_Product | Where-Object {
    $_.Name -match "Java"
}
foreach ($a in $app) {$a.Uninstall()}
 
if (!$app) {Write-Host "Java uninstalled successfully"
Write-Host -ForegroundColor Green "Jeg har nu Afinstalleret Java for: $GivemeName`nUdover det, har jeg installeret den nye version!."
}
if ($app)  {Write-Host  -ForegroundColor Red "Java er nu afinstalleret for: $GivemeName`nGenstart Powershell scriptet for at Installere den nye Java Version!"
}

sleep 10
if(!$app -eq "Java uninstalled successfully"){


$folderOne = "\\$($computer.name)\C$\testfolder\"
if((Test-Path -Path "\\$($computer.name)\C$\testfolder\" -pathtype Container) -eq $False){
Write-Host -ForegroundColor Green "Opretter Mappe: $folderOne ...!"
        New-Item -ItemType directory -Path $folderOne
        }else{
       Write-Host -ForegroundColor Cyan "Du har allerede mappen $folderOne...!"
        }
 
 
 ########################  To get all software installed on the computer
 # Get-WmiObject -Class Win32_Product -ComputerName <changethis> |Select IdentifyingNumber,Name,Version| FT -AutoSize

 ##################### TO get the version of Java:
 # gwmi win32_product -ComputerName <changethis> -filter "Name LIKE '%Java%'"
 ########################


########################################   TO Uninstall Java based on version number! do this  

# $classKey="IdentifyingNumber=`"`{<changethis>`}`",Name=`"<changethis>`",version=`"<changethis>`""
# 
# foreach($computers in $computernames){ ([wmi]"\\$($computers.Name)\root\cimv2:Win32_Product.$classKey").Uninstall() }

###############################################
##### If this does not work use PSEXEC with run as Administrator | This does Install the newest java version!

# $computers = Get-ADComputer -Filter * -SearchBase "OU=Computers,OU=domain,DC=domain,DC=local" | Select-Object Name
# foreach($computernames in $computers){
# \\COMPUTERNAME\C$\PSTools\PsExec.exe \\$(computernames.Name) -u domain\administrator -p PASSWORD msiexec /i "\\SERVER.domain.local\IXP$\SW\JAV00009\IMG\jre1.7.0_67.msi" /qn /norestart /log "\\FILESERVER01\IT\Dokumentation\JavaVersions\Setup_$($computernames.Name).log"
# }
##############    

MSIEXEC /I "\\SERVER.domain.local\IXP$\SW\JAV00009\IMG\jre1.7.0_67.msi" /qn /norestart /log "\\FILESERVER01\IT\Dokumentation\JavaVersions\Setup_$GivemeName.log"
Write-Host -ForegroundColor Green "Jeg har nu oprettet en log for: $GivemeName`n"
reg add HKLM\software\javasoft /v "SPONSORS" /t REG_SZ /d "DISABLE" /f  # No advertisements(ASK toolbar)
reg add HKLM\SOFTWARE\Wow6432Node\JavaSoft /v "SPONSORS" /t REG_SZ /d "DISABLE" /f # No advertisements(ASK toolbar)
Write-Host -ForegroundColor Green "`nJeg har nu fjerne Ask toolbar, som default følger med installationen."

}
}
Exit-PSSession
$Logfolder = " \\FILESERVER01\IT\Dokumentation\JavaVersions\"
Write-Host -ForegroundColor Yellow "Du kan finde alle logs på: $Logfolder`n"