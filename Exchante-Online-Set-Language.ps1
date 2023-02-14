##########
# Set M365 language to Swiss German
# 
# Requirements:
# ExchangeOnlineManagement module:
# https://www.powershellgallery.com/packages/ExchangeOnlineManagement/
##########



# Variables
$admin = <admin@contoso.onmicrosoft.com>

# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName $admin

# Quietly set execution policy to get full execution rights for follwing commands
Set-ExecutionPolicy Bypass -scope Process -Force -Confirm:$False

# Get all users in Exchange Online
$Users = Get-Mailbox -ResultSize Unlimited

# List all mailboxes before changing the language and date/time settings
Write-Host "Mailboxes before changes:"
foreach ($User in $Users) {
  Get-MailboxRegionalConfiguration -Identity $User.Alias | Select Identity,Language,DateFormat,TimeFormat,TimeZone
}

# Loop through each user and change the language and date/time settings
foreach ($User in $Users) {
  #Set-Mailbox -Identity $User.Alias -PreferredLanguage "de-CH" and adapt the mailbox folder names to the chosen language
  Get-MailboxRegionalConfiguration -Identity $User.Alias | Set-MailboxRegionalConfiguration -Language de-CH -DateFormat "dd.MM.yyyy" -TimeFormat "HH:mm" -TimeZone "W. Europe Standard Time" -LocalizeDefaultFolderName:$true
}

# List all mailboxes after changing the language and date/time settings
Write-Host "Mailboxes after changes:"
foreach ($User in $Users) {
  Get-MailboxRegionalConfiguration -Identity $User.Alias | Select Identity,Language,DateFormat,TimeFormat,TimeZone
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline
