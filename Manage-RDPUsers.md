🛡️ PowerShell RDP Manager v5.1

Ein Werkzeug zur Verwaltung von lokalen Remotedesktop-Berechtigungen (RDP) und zur Verteilung von Zugangsdaten per E-Mail.

⚠️ WICHTIGE INFO: Admin vs. Benutzer

Microsoft Outlook erlaubt keinen Zugriff durch Programme, die "Als Administrator" ausgeführt werden. Die Verwaltung von Rechten auf entfernten PCs erfordert jedoch Administrator-Rechte.

Die Lösung: Das Skript trennt diese Aufgaben oder nutzt SMTP.

📋 Beispiele & Anwendungsfälle

Fall 1: Standard (Outlook vorhanden, kein SMTP)

Sie möchten Rechte setzen und E-Mails via Outlook senden.

Schritt 1: Starten Sie PowerShell Als Administrator.

.\Manage-RDPUsers.ps1 -SetRDPRights -UserListPath "users.csv" -ClientListPath "clients.csv" -SendEmail


Ergebnis: Rechte werden gesetzt. Das Skript warnt, dass Outlook nicht als Admin geht, und erstellt ein sendMails...ps1 im Ausgabeordner.

Schritt 2: Starten Sie PowerShell Als Normaler Benutzer (nicht Admin).

& ".\Logs\RDP_Output_...\sendMails_....ps1"


Ergebnis: Outlook öffnet sich und versendet die Mails.

Fall 2: SMTP-Versand (Alles in einem)

Sie haben einen internen Mail-Server (z.B. Exchange), der E-Mails annimmt.

Schritt 1: Starten Sie PowerShell Als Administrator.

.\Manage-RDPUsers.ps1 -SetRDPRights -UserListPath "users.csv" -ClientListPath "clients.csv" -SendEmail -SmtpServer "smtp.firma.local"


Ergebnis: Rechte werden gesetzt UND E-Mails werden direkt versendet (Credentials werden abgefragt).

Fall 3: Nur Rechte entfernen

Sie möchten die Rechte wieder entziehen.

Schritt 1: Starten Sie PowerShell Als Administrator.

.\Manage-RDPUsers.ps1 -RemoveRDPRights -UserListPath "users.csv" -ClientListPath "clients.csv"
