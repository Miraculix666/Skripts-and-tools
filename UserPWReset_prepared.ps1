# Pre-filled options
$preFilledUserName = ""
$preFilledWildcard = "DWR45*"
$preFilledOU = "OU=47678"
$skipConfirmation = $false

# Prompt for the password securely
$securePassword = Read-Host "Bitte geben Sie das Passwort ein" -AsSecureString
$preFilledPassword = [System.Net.NetworkCredential]::new("", $securePassword).Password

# Run the primary script with pre-filled options
.\PrimaryScript.ps1 -UserName $preFilledUserName -Wildcard $preFilledWildcard -OU $preFilledOU -Password $preFilledPassword -SkipConfirmation:$skipConfirmation

# Clean up plaintext password from memory
Remove-Variable preFilledPassword -ErrorAction SilentlyContinue
