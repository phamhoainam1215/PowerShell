# Author:	Pham Hoai Nam <phamhoainam1215@gmail.com>
#
# Date created:	25/Apr/2017
#
# Purpose: Insert auto reply message for users with Auto Reply State is Disabled
#		1. Open script by notepad and edit:
#			-	Username and Password admin(domain).
#			-	Auto reply message
#		2. Run script or schedule it.
#
#About me: https://phamhoainam1215.wordpress.com/ | http://www.office365vietnam.info/ | https://github.com/phamhoainam1215


Import-Module MSOnline

#input username, password and tenant
$username = "user@abc.com"
$password = ConvertTo-SecureString "password_here" -AsPlainText -Force

#input your OOOMessage
$OOOMessage = "Thank you for your email. Your message is important to me and I will respond as soon as possible. Thank You!"

#Connect-MsolService
$UserCredential = New-Object System.Management.Automation.PSCredential $username, $password
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri  https://outlook.office365.com/powershell -Credential $UserCredential -Authentication Basic -AllowRedirection

#Import-PSSession
Import-PSSession $Session -AllowClobber
Connect-MsolService -Credential $UserCredential

#Get all user with AutoReplyState is Disabled
$UserOOO = Get-mailbox -ResultSize Unlimited | Get-MailboxAutoReplyConfiguration | Where-Object { $_.AutoReplyState -eq "Disabled" } | Select Identity,MailboxOwnerId,StartTime,EndTime,AutoReplyState,InternalMessage,ExternalMessage
ForEach ($User in $UserOOO)
{
	$userID = $User.MailboxOwnerId+"@abc.com"
	#set OOOMessage
	Set-MailboxAutoReplyConfiguration -Identity $userID -InternalMessage $OOOMessage -ExternalMessage $OOOMessage
}
Get-PSSession | Remove-PSSession

