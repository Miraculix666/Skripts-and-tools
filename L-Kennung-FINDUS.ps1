# CSV-Datei zum Speichern der Ergebnisse erstellen
$csvPath = "C:\AD_User_FINDUS_Gruppen.csv"

# Kopfzeile in die CSV-Datei schreiben
"Gruppenname,Benutzername" | Out-File -FilePath $csvPath -Encoding UTF8

# Gruppen finden, die "FINDUS" im Namen enthalten
$gruppen = Get-ADGroup -Filter "Name -like '*FINDUS*'" -Properties Name, Members

foreach ($gruppe in $gruppen) {
  foreach ($mitglied in $gruppe.Members) {
    try {
      $user = Get-ADUser -Identity $mitglied -Properties Name, SamAccountName
      if ($user.Name -like "L*") {
        # Ausgabe der relevanten Informationen in die CSV-Datei schreiben
        "$($gruppe.Name),$($user.SamAccountName)" | Out-File -FilePath $csvPath -Encoding UTF8 -Append
      }
    } catch {
      # Fehlerbehandlung, falls das Mitglied keine AD-Benutzer ist
      Write-Warning "Mitglied '$mitglied' in Gruppe '$($gruppe.Name)' konnte nicht als Benutzer aufgelöst werden."
    }
  }
}

# Ausgabe einer Erfolgsmeldung
Write-Host "Die Daten wurden erfolgreich in '$csvPath' exportiert."
