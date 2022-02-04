<#
.SYNOPSIS
	This script collects data on Exchange Online mailboxes and returns the last known time stamp of user activity and logon. 
.DESCRIPTION
	This script depends on ExchangeOnline and will attempt to install the module if it is not found.
    All mailboxes in the tenant are assumed.   
	Checks the last logon time and last interaction time of each mailbox.
    Exports report to environment's desktop as a .csv
.EXAMPLE
	.\PS-EXCHONLINE-User-Activity
.NOTES
    Author:             Luke Von Hagel (lvonhagel@brashbit.io)

	License: 			This script is distributed under "THE BEER-WARE LICENSE" (Revision 42):
						As long as you retain this notice you can do whatever you want with this stuff.
						If we meet some day, and you think this stuff is worth it, you can buy me a beer in return.
#>
#Region InstallEXO
try {
    If (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {Install-Module ExchangeOnlineManagement -scope CurrentUser -Force}
}
catch {
    Write-Error "Failed installing module! - ExchangeOnlineManagement"
    {exit}
}
#EndRegion InstallEXO

#Region Main
#List UPNs from Get-mailbox and using them to identify each user's timestamps in Get-MailboxStatistics. This function doesn't appear to work with bulk users otherwise.
Connect-ExchangeOnline -ShowBanner:$false 
$Output = Get-Mailbox | Select-Object userprincipalname | ForEach-Object{ 
    Get-MailboxStatistics -Identity $_.userprincipalname | Select-Object @{n='User';e={$_.DisplayName}}, @{n='Last Logon Time';e={$_.LastLogonTime}}, @{n='Last Interaction Time';e={$_.LastInteractionTime}}
}
$Output | Export-Csv -path $env:USERPROFILE\Desktop\Report.csv -NoTypeInformation
#EndRegion Main

#Region Housekeeping
#Remove used PS session into 365. !!!WARNING!!! Will remove any and all connections with connection config name of Microsoft Exchange.
Get-PSSession | Where-Object -property ConfigurationName -EQ 'Microsoft.Exchange' | Remove-PSSession
#EndRegion Housekeeping