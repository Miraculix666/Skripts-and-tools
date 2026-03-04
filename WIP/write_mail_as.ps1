#### Transportation Rule für Exchange Server um immer mit eine anderen Email Adresse zu schreiben
#### alternativ über Exchange Admin Center (EAC) konfigurieren


# Verbindung zum Exchange Server herstellen
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://your-exchange-server/PowerShell/ -Authentication Kerberos
Import-PSSession $Session

# Variablen definieren
$RuleName = "Absenderadresse für zweites Postfach ändern"
$SecondMailbox = "zweites.postfach@ihrefirma.de"
$PersonalMailbox = "ihr.name@ihrefirma.de"

# Transportregel erstellen
New-TransportRule -Name $RuleName -FromScope "InOrganization" -From $SecondMailbox -SetHeaderName "X-MS-Exchange-Organization-AuthAs" -SetHeaderValue "Internal" -SetFrom $PersonalMailbox

# Verbindung trennen
Remove-PSSession $Session
