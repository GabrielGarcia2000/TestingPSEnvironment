#***********************************************************************
# PowerShell : TestEmail.ps1                                           *
#   Function : Test e-mali message SMPT relay server.                  *
#            :                                                         *
#***********************************************************************
#                 M O D I F I C A T I O N S                            *
# -- Date -- ---- Name ---- --------- Description -------------------- *
# 11/06/2009 Gabriel Garcia Created.                                   *
#                                                                      *
#***********************************************************************

###
### Example 
### PS C:\> C:\Scripts\TestEmail.ps1
###

### Host Name
$strIPGlobalProperties = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()

### Parameters
$emailfrom = [string]$strIPGlobalProperties.HostName + "@domainname.org"
$emailto = "myemail@domainname.org"
$emailsubject = "My subject"
$emailbody = "My email message."

$SMTPServer = "smtp.domainname.org"
$strOutFileName = "C:\Scripts\TestEnvironment\TestEmail_AttchmentSample.txt"
# $SMTPAuthUsername = "MyUserName"
# $SMTPAuthPassword = "MySMTPUserPassword"

###########################
### e-mail output files ###
###########################

$mailmessage = New-Object system.net.mail.mailmessage
$mailmessage.from = ($emailfrom)
$mailmessage.To.add($emailto)
$mailmessage.Subject = $emailsubject
$mailmessage.Body = $emailbody

#############################
### Attach output file(s) ###
#############################

###  1) SQL Instances Inventory
$attachment = New-Object System.Net.Mail.Attachment($strOutFileName, 'text/plain')
$mailmessage.Attachments.Add($attachment)

### $mailmessage.IsBodyHTML = $true
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25) 
### $SMTPClient.Credentials = New-Object System.Net.NetworkCredential("$SMTPAuthUsername", "$SMTPAuthPassword")
$SMTPClient.Send($mailmessage)
