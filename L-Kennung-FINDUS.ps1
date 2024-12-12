# CSV-Datei zum Speichern der Ergebnisse erstellen
$csvPath = "C:\Daten\AD_User_FINDUS_Gruppen.csv"

# Kopfzeile in die CSV-Datei schreiben
"Gruppenname,Benutzername" | Out-File -FilePath $csvPath -Encoding UTF8

# Alle Benutzer mit relevanten Eigenschaften abrufen
$alleBenutzer = Get-ADUser -Filter * -Properties Name, SamAccountName, MemberOf

# Alle Gruppen mit relevanten Eigenschaften abrufen
$alleGruppen = Get-ADGroup -Filter "Name -like '*FINDUS*'" -Properties Name, Members

# Gruppen durchlaufen
foreach ($gruppe in $alleGruppen) {
  # Mitglieder der Gruppe durchlaufen
  foreach ($mitglied in $gruppe.Members) {
    # Benutzer in der Liste aller Benutzer suchen
    $user = $alleBenutzer | Where-Object { $_.DistinguishedName -eq $mitglied }
    if ($user -and $user.Name -like "L*") {
      # Ausgabe der relevanten Informationen in die CSV-Datei schreiben
      "$($gruppe.Name),$($user.SamAccountName)" | Out-File -FilePath $csvPath -Encoding UTF8 -Append
    }
  }
}

# Ausgabe einer Erfolgsmeldung
Write-Host "Die Daten wurden erfolgreich in '$csvPath' exportiert."
