#Purpose of program: Automate common users tasks in Active Directory with user input
#Created: 01-10-14
#Developer: Max Jensen - JET TIME A/S


# Udtrækker brugere fra gruppe + exporter til CSV
####################################
# $Groups = Get-ADGroup -filter {Name -like "Management"} | Select-Object Name #Change This
# ForEach ($Group in $Groups) {Get-ADGroupMember -identity $($group.name) -recursive | Get-ADUser -Properties Name,SamAccountName,Mail  | Select-Object Name,SamAccountName,Mail | export-csv $path -Delimiter ';' -NoType -Encoding UTF8}
# $path = "c:\testfolder\" + $($group.Name) + ".csv"


if (-not (Get-Module ActiveDirectory)){
    Import-Module ActiveDirectory -ErrorAction Stop           
}
 
# set new default password

 Write-Host "`r`n"
 Write-Host -ForegroundColor Cyan "-------------------------------------------------`r`n"
  Write-Host -ForegroundColor Green "Dette program hjælper med daglige opgaver for Active Directory hvis du har en liste af brugere pr linje`r`n"
  Write-Host -ForegroundColor Red "`r1) Bruge default listen: \\FILESERVER01\IT\Udvikling\Powershell\CommonADUserTasks\usersmustnotChangePwd.txt`n"
   Write-Host -ForegroundColor Red "2) Vælg din egen liste med brugere`n"
   Write-Host -ForegroundColor Red "3) Continue to new functions: EXTRACT MEMBERS / EXPORT members / DO both"
  Write-Host -ForegroundColor Red  "4) Hvilken bruger vil du benytte til at sætte expire date på?"
  Write-Host -ForegroundColor Red "5) Find hvilke grupper Brugeren er medlem af"
  Write-Host -ForegroundColor Red "6) Udtræk Access Control List (Bruger rettigheder) fra UNC-share"
   Write-Host -ForegroundColor Red "7) EXIT`r`n"
  Write-Host "`r`n"

 $listOptions = Read-Host "Hvilken mulighed vil du bruge af overstående?"

if($listOptions -eq "1"){ ## Use default list
$userlist = "\\FILESERVER01\IT\Udvikling\Powershell\CommonADUserTasks\usersmustnotChangePwd.txt"
}
    if($listOptions -eq "2"){ ## Choose your own
    $userlist = Read-Host "Vær venlig at angiv en liste med brugernavne i en mappe!"
    }
    if($listOptions -eq "3"){ ## Choose your own
    
    Write-Host -ForegroundColor Cyan "`r`n--------------------------------------------`r`n"
Write-Host "`r`n"
Write-Host -ForegroundColor Green "Følgende muligheder er tilgængelige`r`n"
Write-Host -ForegroundColor Cyan "`r`n--------------------------------------------`r`n"
Write-Host -ForegroundColor Red "1) Find alle medlemmer i guppen, og output til konsollen"
Write-Host -ForegroundColor Red "2) Exporter alle medlemmer i Gruppen til $path"
Write-Host -ForegroundColor Red "3) Gør begge ting"
Write-Host -ForegroundColor Red "4) Find hvilke grupper brugeren er medlem af"
Write-Host -ForegroundColor Cyan "--------------------------------------------`r`n"
$Option = Read-Host "Hvilken mulighed vil du benytte af overstående?"
$OptionGroup = Read-Host "Hvilken gruppe skal vi finde medlemmer i?"
$Groups = Get-ADGroup -filter {Name -like $OptionGroup } | Select-Object Name #Change This
$path = "c:\testfolder\" + $($group.Name) + ".csv" #Change This / Make a folder your own

$pathForACL = Read-Host "Hvilket share vil du lave udtræk fra?"

if($Option -eq "1"){
Write-Host -ForegroundColor Yellow "`r`nThe Result of Members in AD Group`r`n"
Write-Host -ForegroundColor Cyan "`r`n--------------------------------------------`r`n"
ForEach ($Group in $Groups) {
 Get-ADGroupMember -identity $($group.name) -recursive | Get-ADUser -Properties Name,SamAccountName,Mail  | FT -AutoSize Name,SamAccountName,Mail

 Write-Host -ForegroundColor Cyan "`r`n--------------------------------------------`r`n"
    }

}

if($Option -eq "2"){
ForEach ($Group in $Groups) {
Get-ADGroupMember -identity $($group.name) -recursive | Get-ADUser -Properties Name,SamAccountName,Mail  | Select-Object Name,SamAccountName,Mail  |export-csv $path -Delimiter ';' -NoType -Encoding UTF8
Write-Host "Vær venlig at tjekke mappen: $path"
}
}


if($Option -eq "3"){
$path = "c:\testfolder\" + $($group.Name) + ".csv" #Change This / Make a folder your own
ForEach ($Group in $Groups) {
 Get-ADGroupMember -identity $($group.name) -recursive | Get-ADUser -Properties Name,SamAccountName,Mail  | FT -AutoSize Name,SamAccountName,Mail
 Write-Host -ForegroundColor Cyan "`r`n--------------------------------------------`r`n"

Get-ADGroupMember -identity $($group.name) -recursive | Get-ADUser -Properties Name,SamAccountName,Mail  | Select-Object Name,SamAccountName,Mail  |export-csv $path -Delimiter ';' -NoType -Encoding UTF8
       }

Write-Host "`r`nVær venlig at tjekke mappen: $path"
}

}

if($listOptions -eq "4"){
             Write-Host "`r`n"
         Write-Host -ForegroundColor Cyan "-------------------------------------------------`r`n"
         $Username = Read-Host "Hvilken bruger vil du benytte til at sætte expire date på?"
            Write-Host "`r`n"
         Write-Host -ForegroundColor Red "-------------------------------------------------`r`n"
           Write-Host -ForegroundColor Yellow "Example of setting date: MM/DD/YYYY`n`r"
   [datetime]$Date = read-host "Sæt dato for brugere til at expire"
$newexpdate = $Date.AddDays(1) ## Do this because it automatically subtract one day
 #$Date = Read-Host "Sæt dato for brugere til at expire"
 # loop through the list
ForEach ($u in $Username) {
 
    if ( -not (Get-ADUser -LDAPFilter "(sAMAccountName=$u)")) {
        Write-Host "Can't find $u"
    }
    else {
        $user = Get-ADUser -Identity $u
        $user |Set-ADUser -AccountExpirationDate $newexpdate
        Write-Host -ForegroundColor Yellow "`rYou have changed expire date to: $Date on $u`n"
    }
}
         
         }



         if($listOptions -eq "5"){
$useThisUser = Read-Host "Hvilken bruger vil du benytte? "
Write-Host -ForegroundColor Cyan "Brugeren er medlem af:`r`n"
Get-ADPrincipalGroupMembership $useThisUser | select name
} 
    ### No more stuff to do!? - EXIT
         if($listOptions -eq "6"){
            Write-Host "`r`n"
         Write-Host -ForegroundColor Cyan "-------------------------------------------------`r`n"
         Write-Host -ForegroundColor Red "Thank you for using this system to handle users`n`rWe're now exiting, no more stuff to do!"
         exit
         } 

   
# get list of account names (1 per line)
Write-Host "`r`n"
         Write-Host -ForegroundColor Cyan "-------------------------------------------------`r`n"
         Write-Host "Continue program!...`r`n"

  Write-Host -ForegroundColor Green "Dette program hjælper med daglige opgaver for Active Directory hvis du har en liste af brugere pr linje`r`n"
   Write-Host -ForegroundColor Red "1) Vælg din egen liste med brugere`n"
   Write-Host -ForegroundColor Red "`r2) Bruge default listen: \\FILESERVER01\IT\Udvikling\Powershell\CommonADUserTasks\usersmustnotChangePwd.txt`n"

$myNewestOption = Read-Host "Vælg en af følgende muligheder overstående...!"
if($myNewestOption -eq "1"){
$userlist = Read-Host "Vær venlig at angiv en liste med brugernavne i en mappe!"
$list = Get-Content -Path $userlist
}if($myNewestOption -eq "2"){
$list = Get-Content -Path "\\FILESERVER01\IT\Udvikling\Powershell\CommonADUserTasks\usersmustnotChangePwd.txt"
}

Write-Host -ForegroundColor Red "1) Set Users from list to Not change pwd at logon"
Write-Host -ForegroundColor Red "2) Reset users password from list + set users not to change pwd on next logon"
Write-Host -ForegroundColor Red "3) Set users from list to expire on a specific date"
Write-Host -ForegroundColor Red "4) Find users from list if they exsists in Active Directory, else give error"
Write-Host -ForegroundColor Red "5) Continue to new functions: EXTRACT MEMBERS / EXPORT members / DO both"
Write-Host -ForegroundColor Red "6) EXIT"
 Write-Host "`r`n"
 Write-Host -ForegroundColor Cyan "-------------------------------------------------`r`n"
 
$Choose = Read-Host "Hvilke muligheder af overstående vil du benytte?"


 #####  users may not change pwd at logon
 if($Choose -eq "1"){

  Write-Host "`r`n"
 Write-Host -ForegroundColor Cyan "-------------------------------------------------`r`n"
 # loop through the list
ForEach ($u in $list) {
 
    if ( -not (Get-ADUser -LDAPFilter "(sAMAccountName=$u)")) {
        Write-Host "Can't find $u"
    }
    else {
        $user = Get-ADUser -Identity $u
        $user | Set-AdUser -ChangePasswordAtLogon $false
        Write-Host "User must not change password: $u`n"
        
    }
}
 Write-Host -ForegroundColor Cyan "-------------------------------------------------`r`n"
 


 ##### Reset users password from list + set users not to change pwd on next logon
 }if($Choose -eq "2"){

   Write-Host "`r`n"
 Write-Host -ForegroundColor Cyan "-------------------------------------------------`r`n"
 $password = ConvertTo-SecureString -AsPlainText "PASSWORD" -Force 
 # loop through the list
ForEach ($u in $list) {
 
    if ( -not (Get-ADUser -LDAPFilter "(sAMAccountName=$u)")) {
        Write-Host "Can't find $u"
    }
    else {
        $user = Get-ADUser -Identity $u
       $user | Set-ADAccountPassword -NewPassword $password -Reset -Confirm:$false
        $user | Set-AdUser -ChangePasswordAtLogon $false
        Write-Host -ForegroundColor Yellow "We have now changed the password to the default in Our Organisation & $u must not change password at logon!`n"
     
    }
    
}
 Write-Host -ForegroundColor Cyan "-------------------------------------------------`r`n"
 
 } #####  Set users to expire on a specific date
 if($Choose -eq "3"){
   Write-Host "`r`n"
 Write-Host -ForegroundColor Cyan "-------------------------------------------------`r`n"
   Write-Host -ForegroundColor Yellow "Example of setting date: MM/DD/YYYY`n`r"
   [datetime]$Date = read-host "Sæt dato for brugere til at expire"
$newexpdate = $Date.AddDays(1) ## Do this because it automatically subtract one day
 #$Date = Read-Host "Sæt dato for brugere til at expire"
 # loop through the list
ForEach ($u in $list) {
 
    if ( -not (Get-ADUser -LDAPFilter "(sAMAccountName=$u)")) {
        Write-Host "Can't find $u"
    }
    else {
        $user = Get-ADUser -Identity $u
        $user |Set-ADUser -AccountExpirationDate $newexpdate
        Write-Host -ForegroundColor Yellow "`rYou have changed expire date to: $Date on $u`n"
    }
}

 
 ###### Find users if they exsists in AD else give error
 }if($Choose -eq "4"){
 # loop through the list
ForEach ($u in $list) {
 
    if ( -not (Get-ADUser -LDAPFilter "(sAMAccountName=$u)")) {
        Write-Host -ForegroundColor Red "Cannot find users in Active Directory!:"
        Write-Host "$u`n"
    }
    else {$user = Get-ADUser -Identity $u
     Write-Host -ForegroundColor Green "$u does exists in Active Directory"
 

    }
    }

 }

 ###### EXTRACT MEMBERS / EXPORT members / DO both

 if($Choose -eq "5"){
 
Write-Host -ForegroundColor Cyan "`r`n--------------------------------------------`r`n"
Write-Host "`r`n"
Write-Host -ForegroundColor Green "Følgende muligheder er tilgængelige`r`n"
Write-Host -ForegroundColor Cyan "`r`n--------------------------------------------`r`n"
Write-Host -ForegroundColor Red "1) Find alle medlemmer i guppen, og output til konsollen"
Write-Host -ForegroundColor Red "2) Exporter alle medlemmer i Gruppen til $path"
Write-Host -ForegroundColor Red "3) Gør begge ting"
Write-Host -ForegroundColor Cyan "--------------------------------------------`r`n"
$Option = Read-Host "Hvilken mulighed vil du benytte af overstående?"
$OptionGroup = Read-Host "Hvilken gruppe skal vi finde medlemmer i?"
$Groups = Get-ADGroup -filter {Name -like $OptionGroup } | Select-Object Name #Change This
$path = "c:\testfolder\" + $($group.Name) + ".csv" #Change This / Make a folder your own

if($Option -eq "1"){
Write-Host -ForegroundColor Yellow "`r`nThe Result of Members in AD Group`r`n"
Write-Host -ForegroundColor Cyan "`r`n--------------------------------------------`r`n"
ForEach ($Group in $Groups) {
 Get-ADGroupMember -identity $($group.name) -recursive | Get-ADUser -Properties Name,SamAccountName,Mail  | FT -AutoSize Name,SamAccountName,Mail

 Write-Host -ForegroundColor Cyan "`r`n--------------------------------------------`r`n"
    }

}

if($Option -eq "2"){
ForEach ($Group in $Groups) {
Get-ADGroupMember -identity $($group.name) -recursive | Get-ADUser -Properties Name,SamAccountName,Mail  | Select-Object Name,SamAccountName,Mail  |export-csv $path -Delimiter ';' -NoType -Encoding UTF8
Write-Host "Vær venlig at tjekke mappen: $path"
}
} 

if($Option -eq "3"){
$path = "c:\testfolder\" + $($group.Name) + ".csv" #Change This / Make a folder your own
ForEach ($Group in $Groups) {
 Get-ADGroupMember -identity $($group.name) -recursive | Get-ADUser -Properties Name,SamAccountName,Mail  | FT -AutoSize Name,SamAccountName,Mail
 Write-Host -ForegroundColor Cyan "`r`n--------------------------------------------`r`n"

Get-ADGroupMember -identity $($group.name) -recursive | Get-ADUser -Properties Name,SamAccountName,Mail  | Select-Object Name,SamAccountName,Mail  |export-csv $path -Delimiter ';' -NoType -Encoding UTF8
       }

Write-Host "`r`nVær venlig at tjekke mappen: $path"
}

}
### No more stuff to do!? - EXIT
 if($Choose -eq "5"){
    Write-Host "`r`n"
 Write-Host -ForegroundColor Cyan "-------------------------------------------------`r`n"
 Write-Host -ForegroundColor Red "Thank you for using this system to handle users`n`rWe're now exiting, no more stuff to do!"
 }

