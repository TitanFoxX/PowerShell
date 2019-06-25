# - Load required Snap-in for Microsoft Exchange.

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010

# - Enable and set RemoteMailbox for selected user.

Enable-RemoteMailbox "<full AD name>" -PrimarySmtpAddress "<e-mail address>" -Alias "<SamAccountName>" -RemoteRoutingAddress "<e-mail address>@<company>.mail.onmicrosoft.com"
Set-RemoteMailbox "<SamAccountName>" -EmailAddressPolicyEnabled $False -EmailAddresses @{add='<e-mail address>@<company>.mail.onmicrosoft.com'} 