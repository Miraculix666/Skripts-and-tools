$gruppen = Get-ADGroup -Filter "Name -like '*FINDUS*'" -Properties Name, Members

foreach ($gruppe in $gruppen) {
  foreach ($mitglied in $gruppe.Members) {
    try {
      $user = Get-ADUser -Identity $mitglied -Properties Name, SamAccountName
      if ($user.Name -like "L*") {
        # Ausgabe der relevanten Informationen (z.B. Gruppenname, Benutzername)
        Write-Host "Gruppe: $($gruppe.Name), Benutzer: $($user.SamAccountName)"
      }
    } catch {
      # Fehlerbehandlung, falls das Mitglied keine AD-Benutzer ist
    }
  }
}
