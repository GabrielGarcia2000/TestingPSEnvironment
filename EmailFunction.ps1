function sendMail($mailto, $Subject, $Body, $smtpSrvr, $AttchFile){
	### Host Name
	$strIPGlobalProperties = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()

	### Parameters
	$emailfrom = [string]$strIPGlobalProperties.HostName + "@domain.com"
	$emailto = $mailto
	$emailsubject = $Subject
	$emailbody = $Body
	$SMTPServer = $smtpSrvr
	$strOutFileName = $AttchFile

	$mailmessage = New-Object system.net.mail.mailmessage
	$mailmessage.from = ($emailfrom)
	$mailmessage.To.add($emailto)
	$mailmessage.Subject = $emailsubject
	$mailmessage.Body = $emailbody

	$attachment = New-Object System.Net.Mail.Attachment($strOutFileName, 'text/plain')
	$mailmessage.Attachments.Add($attachment)

	### $mailmessage.IsBodyHTML = $true
	$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25) 
	### $SMTPClient.Credentials = New-Object System.Net.NetworkCredential("$SMTPAuthUsername", "$SMTPAuthPassword")
	$SMTPClient.Send($mailmessage)
}

write-host "Sending file as Email attachment using hotmail.."

# Calling function
sendMail "myemail@domain.com" "Subject" "This is email body" "smtp.domain.com" "C:\Myfolder\AttachMeSample.txt"
write-host "Email Sent"
