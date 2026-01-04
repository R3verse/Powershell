#Purpose of program: Hack Exchange 2010 with custom arbitary extraction of mailboxes without any credentials needed
#Created: 07-02-15
#Developer: Max Jensen - JET TIME A/S


#################### Connecting to MailDB and creating mailboxes + enable them
Write-Host -ForegroundColor Cyan "Connecting to exchangemaildb01 .. Please wait!"
$computerName = "exchangemaildb01"
$ExchangeUsername = "domain\administrator"
$pw = "ADMIN PASS"

# Create Credentials
$securepw = ConvertTo-SecureString $pw -asplaintext -force
$cred = new-object -typename System.Management.Automation.PSCredential -argument $ExchangeUsername, $securepw

# Create and use session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$computerName/PowerShell/ -Credential $cred
Import-PSSession $Session
Enter-PSSession $Session
Start-sleep 3

# Login to exchange server first - and open shell, then type in this
#New-ManagementRoleAssignment -Role "Mailbox Import Export" -User "domain\administrator"
#Get-MailboxExportRequest
# get from date + time
#Get-MessageTrackingLog -server exchange02 -sender test@domain.dk -Start "02-07-2015 06:00" -end "02-09-2015 22:00" | out-string -OutVariable Item

# get all msg
#Get-MessageTrackingLog -Server exchange02 -sender test@domain.dk | out-string -OutVariable Items

############################
cd C:\tmp\
# Reference: http://www.msexchange.org/articles-tutorials/exchange-server-2013/management-administration/managing-pst-import-export-process-exchange-server-2013-part2.html
#
# Hack exchange 2010:##################################################
#                    #Export whole mailbox and open it within outlook! - No further permission needed!


# To get the mailbox:
#New-MailboxExportRequest test@domain.dk -FilePath "\\exchange02\C$\tmp\test.pst" 

#To Delete the pst mailbox again, to not fill space:
Get-MailboxExportRequest | Where { $_.Status –eq 'Completed' } | Remove-MailboxExportRequest –Confirm:$False 

#################################################################################################################


# Get EventID error message from exchange!
# Get-MailboxExportRequest

#For de seneste 2 dage
get-messagetrackinglog -Server exchange02 -start (Get-date).adddays(-1) |?{$_.eventid -eq "FAIL"}| select timestamp,sourcecontext,recipientstatus,eventid,sender,recipients,messagesubject,clientip,serverip,messageid|fl * |out-string -OutVariable Item

#get-messagetrackinglog -Server exchange01 -Start (get-date).addhour(-2) -ResultSize 20 -warningaction 0 | select timestamp,sourcecontext,recipientstatus,eventid,sender,recipients,messagesubject,clientip,serverip,messageid|fl * |out-string -OutVariable Item
#get-messagetrackinglog -Server exchange02 -Start "02-06-2015 02:00" -End  "02-09-2015 10:00" -ResultSize 1000 -warningaction 0 |?{$_.eventid -eq "FAIL"}| select timestamp,sourcecontext,recipientstatus,eventid,sender,recipients,messagesubject,clientip,serverip,messageid|fl * |out-string -OutVariable Item
$emailFrom = "no-reply@domain.dk"
$emailTo = "it@domain.dk"
$Server = "exchange02"
$Title = "Last 2 days FAILED Email Report"
$subject = "$Server - $Title"

$body = @"
EventID: "FAIL" on $Server`n

$Item`r`n
$Items

`r`nMed venlig hilsen`r`nMax Jensen - IT
"@



Start-Sleep 3
Send-MailMessage -To $emailTo -From $emailFrom -Subject $subject -Body $Body -Priority high -SmtpServer "exchange02"
Exit-PSSession $Session
exit