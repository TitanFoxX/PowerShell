# Load required SnapIn for Powershell.
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010

# Get all Exchange DB's first, then send (pipe) them to the Clean-cmdlet.
Get-MailboxDatabase | Clean-MailboxDatabase