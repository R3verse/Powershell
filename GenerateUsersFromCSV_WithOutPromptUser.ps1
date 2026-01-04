#Purpose of program: Create Users in AD, And refer to other script to enable mailboxes on Exchange Server
#Created: 21-09-14
#Developer: Max Jensen - Jet Time A/S

$ErrorActionPreference = "SilentlyContinue" 
add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010 -EA SilentlyContinue
Import-Module ActiveDirectory -EA SilentlyContinue
Set-ExecutionPolicy RemoteSigned


####### Check hvis mapper eksistere, ellers Opret mapper 
$folderOne = "C:\testfolder"
$folderTwo = "C:\testfolder\Testme"
if((Test-Path -Path "C:\testfolder" -pathtype Container) -eq $False){
Write-Host -ForegroundColor "Green" "Opretter Mappe: $folderOne ...!"
        New-Item -ItemType directory -Path $folderOne
        }else{
       Write-Host -ForegroundColor Cyan "Du har allerede mappen $folderOne...!"
      
     }
    if((Test-Path -Path "C:\testfolder\Testme" -pathtype Container) -eq $False){
    New-Item -ItemType directory -Path $folderTwo
        Write-Host -ForegroundColor Green "Opretter Mappe: $folderTwo ...!"
        }else{
        Write-Host -ForegroundColor Cyan "Du har allerede mappen $folderTwo...!"
        
     }

#### Checking if user exists, else exit program!
$User = Get-ADUser -LDAPFilter "(sAMAccountName=$username)" 
If ($User -eq $Null)  
 
    {  

Function ImportCSV ($dir) {
$data = import-csv $dir -delimiter ';'
foreach ($entry in $data)
 {
  CreateUser $entry.FirstName $entry.LastName $entry.Username $entry.Department $entry.PayNo $entry.Base $entry.Description $entry.Title
  $secgroups = ($entry.SecurityGroups).split(",")
  foreach ($g in $secgroups) {
  
    AddSecGroup $entry.Username $g
    AddDistGroup $entry.Username $g
   
  }
  echo "$($entry.Username) has been added to security and distribution groups."
 }

}

Function Get-FileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "CSV Files (*.csv)| *.csv"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName

# *** Entry Point to Script ***

Function CreateUser ($FirstName, $LastName, $username, $department, $PayNo, $Base, $Description, $JobTitle) {

 $location = $Base
 
 if ($location -eq "BLL") {
    $Phone = "+45 0000000"
    $Postalcode = "2770"
    $City = "Billund"
    $Address = "Passagerterminalen 10"
    $HomeDir='\\fileserver01\users\'
 }
 elseif ($location -eq "CPH") {
    $Phone = "+45 0000000"
    $Postalcode = "2770"
    $City = "Dragoer"
    $Address = "Airport"
    $HomeDir='\\fileserver02\users\'
 }elseif($location -eq "FTPSERVERHQ"){
 $Phone = ""
    $Postalcode = "2770"
    $City = "Kastrup"
    $Address = ""
    $HomeDir='\\FTPSERVERHQ\e$\FTPSERVERHQ\groups\'
 }elseif($location -eq "FTPSERVERSouth"){
 $Phone = ""
    $Postalcode = "2770"
    $City = "Kastrup"
    $Address = ""
    $HomeDir='\\FTPSERVERSouth\e$\FTPSERVERSouth\groups\'
 }
 else {
 $location = "HQ" 
    $Phone = "+45 0000000"
    $Postalcode = "2770"
    $City = "Kastrup"
    $Address = "Skoejtevej 27-31"
    $HomeDir='\\fileserver01\users\'
 }
 
###############
# GENERATING RANDOM PASSWORD FOR USER
#
################
function New-Password

{
   Param

   (
       [ValidateRange(5,30)]

       [int]$length = 8,

       [switch]$includeLowerCaseLetters,

       [switch]$includeUpperCaseLetters,

       [switch]$includeNumbers,

       [switch]$includePunctuation,

       [switch]$verbose

    )

    function Test-StringContainsCharacter([string]$string,[string]$characterString)

    {

       $characterFound = $false

       for ($i = 0; $i -lt $characterString.Length; $i++)

       {

           if ($string.Contains($characterString.SubString($i,1)))

           {

              $characterFound = $true

           }

       }

       return $characterFound

    }

    function Get-RandomCharacter([string]$charactersToChooseFrom)

    {

       $offSet = Get-Random -Minimum 0 -Maximum (($charactersToChooseFrom.Length) - 1)

       return $charactersToChooseFrom.SubString($offSet,1)

    }

    [bool]$passwordIsValid = $false

    [string]$lowerCaseLetters = "abcdefghijklmnopqrstuvwxyz"

    [string]$upperCaseLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    [string]$numbers = "0123456789"

    [string]$punctuation = "!@#$%^&*?-=_+()"

    [string]$passwordOptions = ""

    [int]$maximumIterations = 15

    [string]$password = ""

    if ($includeLowerCaseLetters)

    {

        $passwordOptions += $lowerCaseLetters

    }

    if ($includeUpperCaseLetters)

    {

        $passwordOptions += $upperCaseLetters

    }

    if ($includeNumbers)

    {

        $passwordOptions += $numbers

    }

    if ($includePunctuation)

    {

        $passwordOptions += $punctuation

    }

    # checking to see if the utility was called without character set switches

    if ($passwordOptions.Length -eq 0)

    {

       throw "You must select at least one set of characters."

    }

    # Checking to see if the length of password requested is shorter than that needed

    # to fit 1 example of each set.  This should not be be necessary as the attribute

    # of the length parameter restricts the value to a minimum of 5.

    if ($verbose)

    {

       Write-Host $passwordOptions

    }

    $passwordIsValid = $false

    $iterationCount = 0

    while((-not $passwordIsValid) -and ($iterationCount -lt $maximumIterations))

    {

       $passwordIsValid = $true

       $password = ""

       $iterationCount++

       for ($i = 1; $i -le $length; $i++)

       {

           $password += Get-RandomCharacter $passwordOptions

       }

       if ($passwordIsValid -and $includeLowerCaseLetters)

       {

           $passwordIsValid = Test-StringContainsCharacter $password $lowerCaseLetters

       }

       if ($passwordIsValid -and $includeUpperCaseLetters)

       {

           $passwordIsValid = Test-StringContainsCharacter $password $upperCaseLetters

       }

       if ($passwordIsValid -and $includeNumbers)

       {

           $passwordIsValid = Test-StringContainsCharacter $password $numbers

       }

       if ($passwordIsValid -and $includePunctuation)

       {

           $passwordIsValid = Test-StringContainsCharacter $password $punctuation

       }

    }  

    if (-not $passwordIsValid)

    {

       $password = ""

       if ($includeLowerCaseLetters)

       {

           $password += Get-RandomCharacter $lowerCaseLetters

       }

       if ($includeUpperCaseLetters)

       {

           $password += Get-RandomCharacter $upperCaseLetters

       }

       if ($includeNumbers)

       {

           $password += Get-RandomCharacter $numbers

       }

       if ($includePunctuation)

       {

           $password += Get-RandomCharacter $Punctuation

       }

       for ($i = $password.Length + 1; $i -le $length; $i++)

       {

           $password += Get-RandomCharacter $passwordOptions

       }

    }

    return $password

}

$theRandomPass = New-Password -length 10 -includeLowerCaseLetters -includeUpperCaseLetters -includeNumbers -includePunctuation
$myPassWord = $theRandomPass
   


########################################
# END OF PASSWORD GENERATOR
########################################   

# If user is not a part of FTP
 $TempSecure = ConvertTo-SecureString "SUperSecretPassw0rd" -AsPlainText -Force
 #$TempSecure = ConvertTo-SecureString $myPassWord -AsPlainText -Force
 
 $Company = "contoso A/S"
 
 $DisplayName="$Firstname $LastName"
 $UPN=$username+"@contoso.local"
 $Path = 'OU=Powershell users,OU=Users,OU=Contoso,DC=contoso,DC=local'

 $dbno = ($PayNo % 3) + 1
 $TargetDatabase = "exchdb0$dbno"
 $PreferredServer = 'dc01.contoso.local'
 $Country = "Denmark"
 $SAM =  $username.'First name' + "." +  $LastName #Testing 
 
 
 ################ Opretter Velkomstbrev
#Open Word docx, Replace texts in word file, save it
Write-Host -ForegroundColor Green "Replacing variables in DOCX  "

$VelkomstBrevPiloterEngelsk = "EngelskPilotSkabelonVelkomstITmailIntranetCrewConnex.docx"
$VelkomstBrevPiloterDansk = "DanskPilotVelkomstITmailIntranetCrewConnex.docx"
$VelkomstBrevAdministrationPlusRemote = "DanskITmailIntranetRemote.docx"
$MEKSkabelonVelkomstITmailIntranetPlusRemote = "MEKSkabelonVelkomstITmailIntranetPlusRemote.docx"
$DriversHQVelkomstbrev = "DriversHQVelkomstbrev.docx"
$EksternSkabelonVelkomstRemotePlusOases = "EksternSkabelonVelkomstRemotePlusOases.docx"
$FTPSERVERHQUserVelkomstbrevDansk = "FTPSERVERHQUserVelkomstbrevDansk.docx"
$FTPSERVERSouthUsersVelkomst = "FTPSERVERSouthUserVelkomstbrevDansk.docx"

$objPath = "C:\testfolder\Testme\Velkomstbrev IT - $DisplayName.docx"
$objWord = New-Object -ComObject word.application
$objWord.Visible = $False


$objDoc = $objWord.Documents.Open("C:\testfolder\Testme\$MEKSkabelonVelkomstITmailIntranetPlusRemote") #Change this for automate letters


$objSelection = $objWord.Selection

    $FindFirstNameLastName = "FirstLastName"
    $ReplaceFirstNameLastName = "$DisplayName"

    $FindBogstavskode = "bogstavskode"
    $ReplaceBogstavskode = $username

    $FindKodeord = "kodeord"
   #$ReplaceKodeord = $theRandomPass
     $ReplaceKodeord = "pleaseGivemeAPassword"
    
    $FindMailAdress = "MailAdress"
    $ReplaceMailAdress = $username

    $FindCrewUserName = "crewUserName"
    $ReplaceCrewUserName = $username
    
    $FindCrewPassword = "crewPassword"
    $ReplaceCrewPassword = "whateverpassswordIlike"



$ReplaceAll = 2
$FindContinue = 1
$MatchCase = $False
$MatchWholeWord = $True
$MatchWildcards = $False
$MatchSoundsLike = $False
$MatchAllWordForms = $False
$Forward = $True
$Wrap = $FindContinue
$Format = $False

$objSelection.Find.Execute($FindFirstNameLastName,$MatchCase,
  $MatchWholeWord,$MatchWildcards,$MatchSoundsLike,
  $MatchAllWordForms,$Forward,$Wrap,$Format,
  $ReplaceFirstNameLastName,$ReplaceAll)
  
  
  $objSelection.Find.Execute($FindBogstavskode,$MatchCase,
  $MatchWholeWord,$MatchWildcards,$MatchSoundsLike,
  $MatchAllWordForms,$Forward,$Wrap,$Format,
  $ReplaceBogstavskode,$ReplaceAll)
  
  $objSelection.Find.Execute($FindKodeord,$MatchCase,
  $MatchWholeWord,$MatchWildcards,$MatchSoundsLike,
  $MatchAllWordForms,$Forward,$Wrap,$Format,
  $ReplaceKodeord,$ReplaceAll)
  
  $objSelection.Find.Execute($FindMailAdress,$MatchCase,
  $MatchWholeWord,$MatchWildcards,$MatchSoundsLike,
  $MatchAllWordForms,$Forward,$Wrap,$Format,
  $ReplaceMailAdress,$ReplaceAll)
  
  $objSelection.Find.Execute($FindCrewUserName,$MatchCase,
  $MatchWholeWord,$MatchWildcards,$MatchSoundsLike,
  $MatchAllWordForms,$Forward,$Wrap,$Format,
  $ReplaceCrewUserName,$ReplaceAll)
  
  $objSelection.Find.Execute($FindCrewPassword,$MatchCase,
  $MatchWholeWord,$MatchWildcards,$MatchSoundsLike,
  $MatchAllWordForms,$Forward,$Wrap,$Format,
  $ReplaceCrewPassword,$ReplaceAll)

 
$objDoc.SaveAs([REF]$objPath) #Save as $Username docx file 
$objDoc.Close()
$objDoc = $null
$objDoc.Quit()

$timer = 5
$date = Get-Date #Kill that job...!

if (Get-Process winword*) 
    { Get-Process winword | foreach { 
        if((($date - $_.StartTime).seconds) -gt $timer) {
            $procID = $_.id
            Write-Host -ForegroundColor Magenta "Process $procID is running longer than $timer seconds."
            Write-Host -ForegroundColor Green "Killing process $procID.."
            Stop-Process $procID
        }

 } }

 Write-Host -foregroundcolor Green "Velkomstbrev er nu oprettet for : $username"   


##########################
 
  #Check if the User exists  
  $NameID = $user.ID 

 set-adserversettings -preferredserver $PreferredServer
 if($username -like "*FTP_*"){
  new-aduser -path $Path -HomeDirectory "$HomeDir$username" -Name $username -DisplayName $DisplayName -Streetaddress $address -GivenName $FirstName -SurName $LastName -Office $location -AccountPassword $TempSecure -samaccountname $username -PostalCode $Postalcode -EmailAddress "$username@contoso.dk" -ScriptPath "SBS_LOGIN_SCRIPT.bat" -HomeDrive "U:" -UserPrincipalName $UPN -Enabled $true -city $City -company $Company -department $Department -mobilephone $Phone -title $JobTitle -description $Description
 }else{
 new-aduser -path $Path -HomeDirectory "$HomeDir$username" -Name $DisplayName -DisplayName $DisplayName -Streetaddress $address -GivenName $FirstName -SurName $LastName -Office $location -AccountPassword $TempSecure -samaccountname $username -PostalCode $Postalcode -EmailAddress "$username@contoso.dk" -ScriptPath "SBS_LOGIN_SCRIPT.bat" -HomeDrive "U:" -UserPrincipalName $UPN -Enabled $true -city $City -company $Company -department $Department -mobilephone $Phone -title $JobTitle -description $Description -Server DC02
 
$username | Set-AdUser -ChangePasswordAtLogon $false # User must not change password - set to false
$username | Set-ADUser -PasswordNeverExpires $false # User may not have password to never expire
echo "User $username, Must NOT change password at logon.`n"
echo "$username May not have password set to never expire!`n"
}else{
    $username | Set-AdUser -ChangePasswordAtLogon $false # User must not change password - set to false
    $username | Set-ADUser -PasswordNeverExpires $true # User may not have password to never expire
}

start-sleep 3
 set-user -Identity $username -Pager $PayNo -Phone $Phone -Country "Denmark"
 Set-ADUser $username -Replace @{pager=$PayNo}
 new-item -path $HomeDir -name $username -type directory | out-null
 $Foldername=$HomeDir+"$username"
$Acl = Get-Acl $Foldername
$Ar = New-Object  system.security.accesscontrol.filesystemaccessrule("contoso\$username","ListDirectory, ReadData, WriteData, CreateFiles, CreateDirectories, AppendData, ReadExtendedAttributes, WriteExtendedAttributes, Traverse, ExecuteFile, ReadAttributes, WriteAttributes, Write, Delete, ReadPermissions, Read, ReadAndExecute, Modify","ObjectInherit,ContainerInherit","None","Allow")
$Acl.SetAccessRule($Ar)
Set-Acl $Foldername $Acl
 echo "Home directory created on $foldername"
 start-sleep 3

$date = get-date -uformat "%Y_%m_%d_%I%M%p"
#Email structure  
 
$smtp = "exchcas01"
$to = "powershell01@contoso.dk"
$from = "no-reply@contoso.dk"
$subject = "New Users Added to AD: $date" 
$body = "<h1>New User Information</h1><br>"
$body += "<p> <b>First Name:</b> $FirstName </p>"
$body += "<p> <b>Last Name:</b> $LastName </p>"
$body += "<p> <b>Mail:</b> $username@contoso.dk </p>"
$body += "<p> <b>Username:</b> $username </p><br>"
$body += "<p> <b>AD account <u>successfully</u> created!</p>"

#### Now send the email using \> Send-MailMessage 
$encoding = [System.Text.Encoding]::UTF8
send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -Priority high -Encoding $encoding
    
      
 }
 # Adding Security Groups + Destribution Groups
 Function AddSecGroup($username, $group) {
    $g = Get-ADGroup $group
    $groupdn = $g.DistinguishedName
    $fqgroup = [ADSI]("LDAP://$groupdn")
    $fqgroup.member.add("$username") | out-null
   $fqgroup.setinfo() | out-null
   
   Add-ADPrincipalGroupMembership -Identity $username -MemberOf $groupdn
 }
 Function AddDistGroup($username, $group) {
Add-DistributionGroupMember -Identity $group -Member $username -EA SilentlyContinue

 }
      
 
$dir = Get-FileName -initialDirectory "C:\testfolder\"
ImportCSV $dir
$timer = 10
$date = Get-Date #Kill that job...!

if (Get-Process winword*) 
    { Get-Process winword | foreach { 
        if((($date - $_.StartTime).seconds) -gt $timer) {
            $procID = $_.id
            Write-Host -ForegroundColor Magenta "Process $procID is running longer than $timer seconds."
            Write-Host -ForegroundColor Green "Killing process $procID.."
            Stop-Process $procID
        }

 } }
#Converting to PDF from DOC/DOCX
#################################
$wdFormatPDF = 17
$word = New-Object -ComObject word.application
$word.visible = $false
$folderpath = "C:\testfolder\Testme\*"
$fileTypes = "*.docx","*doc"

$VelkomstBrevPiloterEngelsk = "EngelskPilotSkabelonVelkomstITmailIntranetCrewConnex.docx"
$VelkomstBrevPiloterDansk = "DanskPilotVelkomstITmailIntranetCrewConnex.docx"
$VelkomstBrevAdministrationPlusRemote = "DanskITmailIntranetRemote.docx"
$MEKSkabelonVelkomstITmailIntranetPlusRemote = "MEKSkabelonVelkomstITmailIntranetPlusRemote.docx"
$EksternSkabelonVelkomstRemotePlusOases = "EksternSkabelonVelkomstRemotePlusOases.docx"
$FTPSERVERHQUserVelkomstbrevDansk = "FTPSERVERHQUserVelkomstbrevDansk.docx"
$FTPSERVERSouthUsersVelkomst = "FTPSERVERSouthUserVelkomstbrevDansk.docx"
$DriversHQVelkomstbrev = "DriversHQVelkomstbrev.docx"

######## Change $variable after -Exclude 
Get-ChildItem -path $folderpath -include $fileTypes -Exclude $VelkomstBrevPiloterEngelsk,$VelkomstBrevPiloterDansk,$MEKSkabelonVelkomstITmailIntranetPlusRemote,$DriversHQVelkomstbrev,$EksternSkabelonVelkomstRemotePlusOases,$FTPSERVERSouthUsersVelkomst,$FTPSERVERHQUserVelkomstbrevDansk,$VelkomstBrevAdministrationPlusRemote|
foreach-object `
{
 $pathTwo =  ($_.fullname).substring(0,($_.FullName).lastindexOf("."))
 Write-Host -ForegroundColor Green "Converting $pathTwo to pdf ..."
 $doc = $word.documents.open($_.fullname)
 $doc.saveas([ref] $pathTwo, [ref]$wdFormatPDF)
 $doc.close()
}
$word.Quit()


#################### Connecting to MailDB and creating mailboxes + enable them
Write-Host -ForegroundColor Cyan "Connecting to exchmaildb01 .. Please wait!"
$computerName = "exchmaildb02"
$ExchangeUsername = "contoso\administrator"
$pw = "AdministratorPassword"

# Create Credentials
$securepw = ConvertTo-SecureString $pw -asplaintext -force
$cred = new-object -typename System.Management.Automation.PSCredential -argument $ExchangeUsername, $securepw

# Create and use session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$computerName/PowerShell/ -Credential $cred
Import-PSSession $Session
Write-Host -ForegroundColor Cyan "Creating mailboxes.. Please wait!"
start-sleep 7
Get-User -OrganizationalUnit "contoso.local/Contoso/Users/Powershell users" -RecipientTypeDetails User | Enable-Mailbox | get-mailbox | select name,windowsemailaddress,database
Get-ADUser -Filter * -SearchBase 'OU=Powershell users,OU=Users,OU=contoso,DC=contoso,DC=local' -Properties userPrincipalName | foreach { Set-ADUser $_ -UserPrincipalName "$($_.samaccountname)@contoso.dk" -Confirm:$false}
Write-Host -foregroundColor Green "You have now enabled all the following mailboxes above!`n"
 Write-Host -foregroundColor Magenta "Please wait while moving users to right OU!.."
Start-Sleep 3



 #### Check if $username is in right OU, else put to default OU
 $DeafultPath = Get-ADUser -Filter * -SearchBase 'OU=Powershell users,OU=Users,OU=contoso,DC=contoso,DC=local'
 $MyDepartment =  Get-ADUser -Filter * -SearchBase 'OU=Powershell users,OU=Users,OU=contoso,DC=contoso,DC=local' -Properties Department
 $MYOU = 'OU=Powershell users,OU=Users,OU=contoso,DC=contoso,DC=local'



$UserDepartments=Get-ADUser -searchbase $Path -filter * -Properties Department| FT -A Department

foreach ($user in $MyDepartment){
 if($user.Department -eq "IT"){ ######## HQ
  Write-Host -ForegroundColor Cyan "We found a IT USER in HQ, Changing OU!..."
  Move-ADObject $user -TargetPath 'OU=Administration,OU=HQ,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local' 
 }elseif($user.Department -eq "Passenger Services"){
  Write-Host -ForegroundColor Cyan "We found a Passenger Services USER in HQ, Changing OU!..."
  Move-ADObject $user -TargetPath 'OU=Administration,OU=HQ,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }elseif($user.Department -eq "Acounting"){
  Write-Host -ForegroundColor Cyan "We found a Acounting USER in HQ, Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=Administration,OU=HQ,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }elseif($user.Department -eq "Sales(HQ)"){
  Write-Host -ForegroundColor Cyan "We found a Sales USER in HQ, Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=Administration,OU=HQ,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }
 elseif($user.Department -eq "Pilot (CPH)"){
  Write-Host -ForegroundColor Cyan "We found a Pilot in CPH , Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=ATR,OU=Pilots,OU=HQ,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }elseif($user.Department -eq "Pilot (BLL)"){ 
  Write-Host -ForegroundColor Cyan "We found a Pilot in BLL , Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=Boeing,OU=Pilots,OU=Billund,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }elseif($user.Department -eq "Pilot (Bergen)"){ 
  Write-Host -ForegroundColor Cyan "We found a Pilot in Bergen , Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=Boeing,OU=Pilots,OU=Billund,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }
 elseif($user.Department -eq "Administration (HQ)"){
 Write-Host -ForegroundColor Cyan "We found a HQ Employee, Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=Administration,OU=HQ,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }elseif($user.Department -eq "OP"){
 Write-Host -ForegroundColor Cyan "We found an OP Employee, Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=Administration,OU=HQ,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }
 
 elseif($user.Department -eq "Engineering"){ ################ Syd
  Write-Host -ForegroundColor Cyan "We found a Engineering in Syd, Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=Administration,OU=Syd,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }elseif($user.Department -eq "Logistics"){
  Write-Host -ForegroundColor Cyan "We found a Logistics in Syd, Changing OU!..."
   Move-ADObject $user -TargetPath 'OU=Administration,OU=Syd,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }elseif($user.Department -eq "Planning"){
  Write-Host -ForegroundColor Cyan "We found a Planning in Syd, Changing OU!..."
  Move-ADObject $user -TargetPath 'OU=Administration,OU=Syd,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }elseif($user.Department -eq "Mekaniker (CPH)"){
  Write-Host -ForegroundColor Cyan "We found a mechanic in CPH, Changing OU!..."
  Move-ADObject $user -TargetPath 'OU=Mechanics,OU=Syd,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }

 elseif($user.Department -eq "Mekaniker (BLL)"){ ################ Billund
 Write-Host -ForegroundColor Cyan "We found a mechanic in BLL, Changing OU!..."
Move-ADObject $user -TargetPath 'OU=Mechanics,OU=Billund,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }
 
 elseif($user.Department -eq "Cabin (HQ)"){ ######## Cabin
  Write-Host -ForegroundColor Cyan "We found a Cabin Attendant User - HQ , Changing OU!..."
  Move-ADObject $user -TargetPath 'OU=Cabin,OU=HQ,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }elseif($user.Department -eq "Cabin (BLL)"){
  Write-Host -ForegroundColor Cyan "We found a Cabin Attendant User - BLL , Changing OU!..."
  Move-ADObject $user -TargetPath 'OU=Cabin,OU=Billund,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
 }
 
 elseif($user.Department -eq "Go Technics"){ ####### Eksterne
  Write-Host -ForegroundColor Cyan "We found an Ekstern Go Technics , Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=Line Stations,OU=Eksterne,OU=Users,OU=contoso,DC=contoso,DC=local'
echo "User: $username, have now changed to password never expires.`n"
$username | Set-ADUser -PasswordNeverExpires $true # User may not have password to never expire

 }elseif($user.Department -eq "Ekstern"){ ####### Eksterne
  Write-Host -ForegroundColor Cyan "We found an Ekstern, Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=Line Stations,OU=Eksterne,OU=Users,OU=contoso,DC=contoso,DC=local'
 echo "User: $username, have now changed to password never expires.`n"
$username | Set-ADUser -PasswordNeverExpires $true # User may not have password to never expire
 }
 elseif($user.Department -eq "Vikar"){
 Write-Host -ForegroundColor Cyan "We found a Vikar, Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=Vikarer,OU=Eksterne,OU=Users,OU=contoso,DC=contoso,DC=local'
 echo "User: $username, have now changed to password never expires.`n"
$username | Set-ADUser -PasswordNeverExpires $true # User may not have password to never expire
 }
 elseif($user.Department -eq "FTPSERVERSouth"){
 Write-Host -ForegroundColor Cyan "We found a FTP user, Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=FTP Brugere,OU=Users,OU=contoso,DC=contoso,DC=local'
 echo "User: $username, have now changed to password never expires.`n"
$username | Set-ADUser -PasswordNeverExpires $true -CannotChangePassword:$true -ChangePasswordAtLogon $false # User may not have password to never expire & May never change password
 } elseif($user.Department -eq "FTPSERVERHQ"){
 Write-Host -ForegroundColor Cyan "We found a FTP user, Changing OU!..."
 Move-ADObject $user -TargetPath 'OU=FTP Brugere,OU=Users,OU=contoso,DC=contoso,DC=local'
 echo "User: $username, have now changed to password never expires.`n"
$username | Set-ADUser -PasswordNeverExpires $true -CannotChangePassword:$true -ChangePasswordAtLogon $false # User may not have password to never expire & May never change password
 }
 elseif($user.Department -eq "OP"){
  Write-Host -ForegroundColor Cyan "We found an Operation USER in HQ, Changing OU!..."
  Move-ADObject $user -TargetPath 'OU=Administration,OU=HQ,OU=Medarbejdere,OU=Users,OU=contoso,DC=contoso,DC=local'
  }
   else{ ###### Default
  Write-Host -ForegroundColor Cyan "I don't know where to put this object, setting to default OU!..."
  Write-Host -ForegroundColor Red "OU=Powershell users,OU=Users,OU=contoso,DC=contoso,DC=local`n"
    
    Move-ADObject $user -TargetPath $Path

    }
}
Remove-PSSession $Session

##### CLEANUP - kill WINWORD
$timer = 5
$date = Get-Date #Kill that job...!

if (Get-Process winword*) 
    { Get-Process winword | foreach { 
        if((($date - $_.StartTime).seconds) -gt $timer) {
            $procID = $_.id
            Write-Host -ForegroundColor Magenta "Process $procID is running longer than $timer seconds."
            Write-Host -ForegroundColor Green "Killing process $procID.."
            Stop-Process $procID
        }

 } }
Write-Host -ForegroundColor Red "`nWe're done here.. Closing session!"
} ElseIF($user -eq $user) {  
      Write-Host -foregroundcolor Red "User Already exist : $username" 
      exit
      }