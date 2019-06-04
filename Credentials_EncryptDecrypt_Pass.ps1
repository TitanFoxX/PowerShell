# Type or copy-paste your plain text password in the pop-up
# The password is then encrypted and stored in the $Encrypted_Pass variable

$Encrypted_Pass = Read-Host -Prompt "Enter your password for encryption" -AsSecureString | ConvertFrom-SecureString
Write-Host "The encrypted password is : $Encrypted_Pass"

# If the application in which you want to use the password doesn't support
# the use of encrypted passwords, this script can be used to decrypt it first

$Unencrypted_Pass = $Encrypted_Pass | ConvertTo-SecureString
$Unencrypted_Pass = [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($Unencrypted_Pass))
Write-Host "The unencrypted password is : $Unencrypted_Pass"