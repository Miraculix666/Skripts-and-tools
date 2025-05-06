
PS C:\Daten\Scripte\wip> C:\Daten\Scripte\AD-Sync-tree.ps1
In C:\Daten\Scripte\AD-Sync-tree.ps1:27 Zeichen:81
+ ... er beim Abrufen der Domänencontroller für Domäne $DomainName: $errorM ...
+                                                      ~~~~~~~~~~~~
Ungültiger Variablenverweis. Nach ":" folgte kein Zeichen, das für einen Variablennamen gültig ist. Verwenden Sie ggf. "${}", 
um den Namen zu begrenzen.
    + CategoryInfo          : ParserError: (:) [], ParentContainsErrorRecordException
    + FullyQualifiedErrorId : InvalidVariableReferenceWithDrive
 




Import-Module ActiveDirectory

try {
    # Aktuellen AD-Forest abrufen


    $Forest = Get-ADForest
    if (-not $Forest) {
        Write-Error "Konnte den Active Directory Forest nicht abrufen. Stellen Sie sicher, dass das Skript auf einer domänenangebundenen Maschine mit installierten RSAT-AD-Tools ausgeführt wird."
        exit 1
    }
    Write-Host "Aktueller Forest: $($Forest.Name)"

    # Alle Domänencontroller im gesamten Forest sammeln
    $AllDCsInForest = @()
    Write-Host "Sammle alle Domänencontroller im Forest..."
    foreach ($DomainName in $Forest.Domains) {
        try {
            Write-Host "Rufe Domänencontroller für Domäne ab: $DomainName"
            $DCsInDomain = Get-ADDomainController -Filter * -Server $DomainName -ErrorAction Stop # -ErrorAction Stop ist gut hier
            if ($DCsInDomain) {
                $AllDCsInForest += $DCsInDomain
            } else {
                Write-Warning "Keine Domänencontroller in der Domäne $DomainName gefunden."
            }
        } catch {
            # KORREKTUR für Zeile ca. 25: Fehlermeldung zuerst in Variable speichern
            $errorMessage = $_.Exception.Message
            Write-Warning "Fehler beim Abrufen der Domänencontroller für Domäne $DomainName: $errorMessage"
        }
    }

    if ($AllDCsInForest.Count -eq 0) {
        Write-Error "Keine Domänencontroller im Forest gefunden oder Zugriff verweigert. Stellen Sie sicher, dass Sie über die erforderlichen Berechtigungen verfügen."
        exit 1
    }

    Write-Host "`nStarte Replikationsanstoß für $($AllDCsInForest.Count) Domänencontroller im Forest..."
    Write-Host "Dies kann einige Zeit dauern, abhängig von der Größe und Topologie Ihres Forests."

    $SuccessCount = 0
    $FailCount = 0
    $FailedDCs = @()

    # Für jeden Domänencontroller im Forest die Replikation anstoßen
    foreach ($DC in $AllDCsInForest) {
        $DCHostname = $DC.HostName
        Write-Host "--------------------------------------------------"
        Write-Host "Stoße Replikation für $DCHostname an..."
        try {
            $Command = "repadmin /syncall $DCHostname /A /e /d"
            Write-Host "Führe aus: $Command"
            
            # Direkter Aufruf von repadmin.exe ist oft besser als Invoke-Expression für externe Befehle
            # und ermöglicht eine einfachere Prüfung von $LASTEXITCODE.
            # Für dieses Beispiel belassen wir Invoke-Expression, um den ursprünglichen Fehler zu beheben.
            Invoke-Expression $Command 
            
            # Hinweis: Invoke-Expression fängt möglicherweise nicht alle Fehler von repadmin ab.
            # Eine robustere Fehlerbehandlung würde 'repadmin.exe' direkt aufrufen und $LASTEXITCODE prüfen.
            # z.B.: & repadmin /syncall $DCHostname /A /e /d
            # if ($LASTEXITCODE -ne 0) { throw "repadmin failed for $DCHostname with exit code $LASTEXITCODE" }


            Write-Host "Replikationsbefehl für $DCHostname gesendet." -ForegroundColor Green
            $SuccessCount++
        } catch {
            # KORREKTUR für Zeile ca. 75: Fehlermeldung zuerst in Variable speichern
            $errorMessage = $_.Exception.Message
            # Die Fehlermeldung hier kommt von Invoke-Expression oder einem Skript-terminierenden Fehler,
            # nicht unbedingt direkt von repadmin, wenn repadmin nur einen non-zero Exit-Code zurückgibt.
            Write-Warning "Fehler beim Anstoßen der Replikation für $DCHostname (während Invoke-Expression): $errorMessage"
            $FailCount++
            $FailedDCs += $DCHostname
        }
    }

    Write-Host "--------------------------------------------------"
    Write-Host "`nZusammenfassung des Replikationsanstoßes:"
    Write-Host "Befehl zum Anstoßen der Replikation erfolgreich gesendet für: $SuccessCount Domänencontroller." -ForegroundColor Green
    if ($FailCount -gt 0) {
        Write-Warning "Fehlgeschlagen für: $FailCount Domänencontroller (oder Fehler im Skriptablauf)."
        Write-Warning "Betroffene DCs (falls Fehler spezifisch für DC-Operation): $($FailedDCs -join ', ')"
        Write-Warning "Überprüfen Sie die Fehlermeldungen oben und stellen Sie sicher, dass 'repadmin.exe' im Pfad ist und Netzwerkverbindung sowie Berechtigungen vorhanden sind."
    }
    Write-Host "Der Replikationsprozess im Hintergrund kann je nach Netzwerk und Auslastung noch einige Zeit in Anspruch nehmen."

} catch {
    # Gesamter try-catch Block
    $ScriptErrorMessage = $_.Exception.Message
    Write-Error "Ein schwerwiegender Fehler im Skript ist aufgetreten: $ScriptErrorMessage"
    Write-Error "Stellen Sie sicher, dass das ActiveDirectory Modul installiert ist (Teil der Remote Server Administration Tools - RSAT) und Sie über ausreichende Berechtigungen verfügen (z.B. Mitglied der Gruppe 'Domänen-Admins' oder 'Organisations-Admins')."
}
