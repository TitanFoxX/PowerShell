# - Display pop-up for filling in your AD-credentials.

$Credentials = Get-Credential

# - Export encrypted credentials to XML-file.

$Credentials | Export-CliXml -Path '<file path>\credentials.xml'