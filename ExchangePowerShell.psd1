@{
RootModule = 'ExchangePowerShell.psm1'
Description ='The EP (ExchangePowerShell) Powershell module is a collection of commandlets that can be used to assist an Exchange Server 2016 administrator to perform common tasks.'
ModuleVersion = '0.11.0'
Author = 'Pietro Ciaccio | LinkedIn: https://www.linkedin.com/in/pietrociaccio | Twitter: @PietroCiac'
FunctionsToExport = @(
'Clear-EPAutoMapping'
'Clear-EPExchangeLogs'
'Convert-EPIMCEAEXtoX500'
'Get-EPMaintenanceMode'
'Disable-EPMaintenanceMode'
'Enable-EPMaintenanceMode'
'Get-EPAntimalwareStatus'
'Import-EPPSTs'
'Remove-EPProxyAddress'
'ConvertTo-EPMailUser'
'ConvertTo-EPMailContact'
'Start-EPRequiredServices'
)
GUID = '722ba3b3-d2cf-4ac4-a6e6-43aeb143835c'
PowerShellVersion = '5.1'
PrivateData = @{
	PSData = @{
		Tags = @('Exchange','PSEdition_Desktop','Windows','Powershell','Server','2016')
		ReleaseNotes = @'
EP was written to support a Microsoft Exchange Server 2016 Organization hosted on Microsoft Windows Server 2016.
Other scenarios may be supported but are untested.
It is the users responsibility to ensure EP is correctly tested before using in a production environment.

# 0.11.0

* Added cmdlet Import-EPPSTs. Used to bulk import multiple PST files into a mailbox or archive mailbox.

# 0.10.1

* Small improvement to Enable-EPMaitenanceMode cmdlet. No longer throws when unable to move an active database copy. Now displays a warning.

# 0.10.0

* Added Get-EPMaintenanceMode cmdlet.
* Added Start-EPRequiredServices cmdlet.
* Small change to maintenance mode cmdlets. Cmdlets no longer throw an exception when unable to restart transport services. A warning is displayed instead.
* Correction in evaluating Exchange admindisplayversion with Exchange Server 2016 cumulative update 16.
* Added UnifiedContent path to Clear-EPExchangeLogs cmdlet. This path is populated but not maintained by Exchange native anti-malware services.

# 0.9.0

* Added Convert-EPIMCEAEXtoX500 cmdlet.

# 0.8.0

* Added Clear-EPAutoMapping cmdlet.

# 0.7.3

* Bug fix with ConvertTo-EPMailContact cmdlet.

# 0.7.2

* Bug fix.

# 0.7.1

* Bug fix.

# 0.7.0

* Added Get-EPAntimalwareStatus cmdlet.

# 0.6.0

* Rebranding of cmdlets from ExPS to EP prefixes.

# 0.5.0

* Removed restriction to Exchange 2016 objects and servers. ExPS is free to be used against any version of Exchange or Operating System.
* Bug fixes with maintenance mode cmdlets.
* Added ConvertTo-ExPSMailContact cmdlet.

# 0.4.0

* Restricted cmdlets to Exchange 2016 objects only.
* Added ConvertTo-ExPSMailUser cmdlet. Use this to convert mailboxes to mail users.

# 0.3.1

* Updated help.
* Bug fix.

# 0.3.0

* Redesign to reduce complexity.
* Removed some cmdlets. These will be added back at a later time.
* Changed to object oriented outputs so they can be stored in variables or written to logs in a structured format.
* Added cmdlet to remove proxy addresses.

# 0.2.3

* Small improvement with loading Exchange cmdlets.
* Bug fix in checking mail object identities.

# 0.2.2

* Bug fix with Get-ExPSMailboxPermission cmdlet. Resolved an issue where multiple mailboxes would be returned from a single query.

# 0.2.1

* Small changes and added switches to Get-ExPSMailboxPermissions cmdlet.

# 0.2.0

* Added Get-ExPSMailboxPermission cmdlet. This will get full access, send as, send on behalf, inbox folder and calendar folder permissions in a single command to provide an overview of who has permissions on a mailbox.

# 0.1.1

* Bug fix with date time issue.

# 0.1.0

* Removed requirement to specify a redirection server with Enable-ExPSMaintenanceMode.
* Removed restriction to Exchange 2016 only. Module will allow you to run against any version of Exchange but will show a warning.
* Added better error handling to cmdlets.
* Added cmdlet Read-ExPSIMAPLogs.
* Added cmdlet Read-ExPSPOPLogs.
* Improvements to date time support.
* Added services retry capability to maintenance mode cmdlets.

# 0.0.0 Initial

* Initial release. 

'@
}
}
}

