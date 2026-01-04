#Purpose of program: Find users who are disabled or expire flag on. Move Disabled users to Disabled OU. Remove all groups from user.
#                    
#Created: 23-09-14
#Developer: Max Jensen - JET TIME A/S

Import-Module Activedirectory

$OU = 'OU=Medarbejdere,OU=Users,OU=domain,DC=domain,DC=local'
$OUTWO = 'OU=Users,OU=domain,DC=domain,DC=local'
Write-Host -ForegroundColor Magenta "Here is the AD accounts for disabled users in: $OU"
Search-AdAccount -AccountDisabled -SearchBase $OU | Format-Table Name,SamAccountName
 Write-Host -ForegroundColor Magenta "`n`rHere is the user accounts who are expiring in: $OUTWO"
#Search-ADAccount -AccountExpiring -TimeSpan "90" -SearchBase $OU | Select-Object Name,SamAccountName,AccountExpirationDate | Sort-Object AccountExpirationDate
Search-ADAccount -AccountExpiring -TimeSpan "90" -SearchBase $OUTWO | Select-Object Name,SamAccountName,AccountExpirationDate | Sort-Object AccountExpirationDate

# Get Dates: today and today+365 days. The second varible can be used to send a notification email to the
# manager field of the account
$Today= Get-Date
$TodayP30 = $Today.AddDays(365)
 
# Define the OU path where user accounts reside and the Suspended OU LDAP path
$SuspendedOUPath = "OU=DisabledUsers(not to be deleted),OU=Users,OU=domain,DC=domain,DC=local"
 
# Searchscope 2 = Subtree
$users=Search-AdAccount -AccountDisabled -SearchBase $OU

# Moving user and remove it from all it's groups

 
FOREACH ($User in $users) { 
    
    Move-ADObject $User -TargetPath $SuspendedOUPath #Moving the target to Disabled users
     Write-Host -ForegroundColor Green "You have now moved:$User to: $SuspendedOUPath"

     Get-ADGroup -Filter 'GroupCategory -eq "Security" -or GroupCategory -eq "Distribution"' -SearchBase $OU | ForEach-Object{ $group = $_
	Get-ADGroupMember -Identity $group -Recursive | %{Get-ADUser -Identity $_.distinguishedName -Properties Enabled | ?{$_.Enabled -eq $false}} | ForEach-Object{ $user = $_
		$uname = $user.Name
		$gname = $group.Name
		Write-Host "Removing $uname from $gname" -Foreground Yellow
		Remove-ADGroupMember -Identity $group -Member $user -Confirm
    Write-Host -ForegroundColor Green "You have now Removed groups from $user" #Removing groups from users

	}
}
} 

# ------------------------------------------------------------------
# Get list of users from Active Directory that will expire in 7 days
# ------------------------------------------------------------------

$UserList = Search-ADAccount -AccountExpiring -UsersOnly -TimeSpan 90.00:00:00 -SearchBase $OUTWO | Select-Object Name,SamAccountName,AccountExpirationDate | Sort-Object AccountExpirationDate
$UserList2 = Search-AdAccount -AccountDisabled -SearchBase $OU |Select-Object Name,SamAccountName | Sort-Object Name 

$date = get-date -uformat "%Y_%m_%d_%I%M%p"
#Email structure  
$ReportName = "C:\Expiring_Users_$date.csv"
$ReportName2 = "C:\Disabled_Users_$date.csv"
 
$smtp = "exchcas01"
$to = "IT@domain.dk"
$from = "no-reply@domain.dk"
$subject = "Weekly expiring & Disabled User Report: $date" 
$body = "<b>Attachment for disabled users:</b> $ReportName2<br>"
$body += "Please note - USERS ARE MOVED TO THIS OU:<br>"
$body += "<b>OU=DisabledUsers(not to be deleted),OU=Users,OU=domain,DC=domain,DC=local</b>"

$body += "<p><b>Attachment for Expired users - Timespan is 90 Days:</b> $ReportName"
$body += "<br><br><b>This is the OU we are looking in:</b><br>OU=Users,OU=domain,DC=domain,DC=local<br>"
$body += "<br><br>Med venlig hilsen, <br><b>Mr. Powershell Guru <._.></b>"



# -----------------------------------------------
# Send an email using the variables defined above
# -----------------------------------------------

If ($UserList -eq $null){}
Else
{
   $UserList | Export-CSV $ReportName
   $UserList2 | Export-CSV $ReportName2
   #### Now send the email using \> Send-MailMessage 
$encoding = [System.Text.Encoding]::UTF8
send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body $body -Attachments $ReportName,$ReportName2 -BodyAsHtml -Priority high -Encoding $encoding

}