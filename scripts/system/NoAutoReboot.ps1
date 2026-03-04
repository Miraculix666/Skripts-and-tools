# Define the path to the registry key
$Path = "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU"

# Create the key path if it doesn't already exist. 
# The -Force parameter creates all parent keys in the path as needed.
New-Item -Path $Path -Force | Out-Null

# Now, set the desired value to prevent auto-reboot
Set-ItemProperty -Path $Path -Name "NoAutoRebootWithLoggedOnUsers" -Value 1

Write-Host "Registry value 'NoAutoRebootWithLoggedOnUsers' has been set successfully."