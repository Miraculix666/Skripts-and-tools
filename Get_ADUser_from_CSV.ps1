# Benutzernamen aus dem AD abrufen und in eine Liste schreiben
$UserList = Get-ADUser -Filter * | Select-Object -ExpandProperty SamAccountName

# Liste der Benutzernamen ausgeben
$UserList | Out-File -FilePath "C:\Pfad\Zur\Datei\Benutzerliste.txt"

# Erfolgsmeldung
Write-Host "Die Benutzerliste wurde erfolgreich erstellt und unter C:\Pfad\Zur\Datei\Benutzerliste.txt gespeichert."
