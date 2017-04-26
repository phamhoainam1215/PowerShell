# Author:	Pham Hoai Nam <phamhoainam1215@gmail.com>
#
# Date created:	25/Apr/2017
#
# Purpose: Insert auto reply message for users with Auto Reply State is Disabled and user with license determined
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

#get tenant: connect-msolservice and Get-MsolAccountSku
$typeOfLicense = "abccom:EXCHANGEENTERPRISE_FACULTY"

#input your OOOMessage
$OOOMessage = 'Thank you for your email. Your message is important to me and I will respond as soon as possible. Thank You!'

#Connect-MsolService
$UserCredential = New-Object System.Management.Automation.PSCredential $username, $password
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri  https://outlook.office365.com/powershell -Credential $UserCredential -Authentication Basic -AllowRedirection

#Import-PSSession
Import-PSSession $Session -AllowClobber
Connect-MsolService -Credential $UserCredential

#Get users with Licenses
$liceneseType = @{n="Licenses Type";e={$_.Licenses.AccountSKUid}}
$proxyUser = @{n="ProxyAddresses";e={$_.ProxyAddresses}}
$UserOOOwithLicenseStudent = Get-MsolUser -All |Where {$_.IsLicensed -eq $true } |Select DisplayName,UsageLocation,$liceneseType,SignInName,UserPrincipalName,$proxyUser
ForEach ($User in $UserOOOwithLicenseStudent)
{
	#Filter user with EXCHANGESTANDARD license
	if( $User.'Licenses Type' -eq $typeOfLicense )
	{
		#Get users with AutoReplyState is Disabled
		$UserOOO = Get-Mailbox -Identity $User.UserPrincipalName | Get-MailboxAutoReplyConfiguration | Where-Object { $_.AutoReplyState -eq "Disabled" } #| Select Identity,MailboxOwnerId,StartTime,EndTime,AutoReplyState,InternalMessage,ExternalMessage
		ForEach ($UserO in $UserOOO)
		{
			$userID = $User.UserPrincipalName
			#set OOOMessage
			Set-MailboxAutoReplyConfiguration -Identity $userID -InternalMessage $OOOMessage -ExternalMessage $OOOMessage$userDisplayName		
		}
	}
}

#Remove-PSSession
Get-PSSession | Remove-PSSession

