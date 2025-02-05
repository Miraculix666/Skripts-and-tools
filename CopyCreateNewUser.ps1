PS C:\Daten\Scripte> .\CopyCreateNewUser.ps1 -TemplateUser L7111101 -NewUserName L7111115
In C:\Daten\Scripte\CopyCreateNewUser.ps1:200 Zeichen:56
+ ... Write-Log "Fehler beim Erstellen des Benutzers $NewSam: $_" -Le ...
+ ~~~~~~~~
Ungültiger Variablenverweis. Nach ":" folgte kein Zeichen, das für einen Variablennamen gültig ist. Verwenden Sie ggf. "${}", um den 
Namen zu begrenzen.
In C:\Daten\Scripte\CopyCreateNewUser.ps1:258 Zeichen:23
+ if ($newUser) {
+ ~
Die schließende "}" fehlt im Anweisungsblock oder der Typdefinition.
In C:\Daten\Scripte\CopyCreateNewUser.ps1:234 Zeichen:28
+ foreach ($u in $users) {
+ ~
Die schließende "}" fehlt im Anweisungsblock oder der Typdefinition.
In C:\Daten\Scripte\CopyCreateNewUser.ps1:211 Zeichen:48
+ if ($PSBoundParameters.ContainsKey("CsvPath")) {
+ ~
Die schließende "}" fehlt im Anweisungsblock oder der Typdefinition.
+ CategoryInfo : ParserError: (:) [], ParseException
+ FullyQualifiedErrorId : InvalidVariableReferenceWithDrive
