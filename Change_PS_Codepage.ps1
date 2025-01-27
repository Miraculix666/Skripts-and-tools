#Somtimes Powershell has a wrong Codepage Setting, this script can correct this


# Setzt die Codepage auf UTF-8
chcp 65001

# Setzt die Umgebungsvariable auf UTF-8
$env:LC_ALL = "de_DE.UTF-8"

# Setzt die Konsolenausgabe auf UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
