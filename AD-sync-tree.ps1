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
            $DCsInDomain = Get-ADDomainController -Filter * -Server $DomainName -ErrorAction Stop
            if ($DCsInDomain) {
                $AllDCsInForest += $DCsInDomain
            } else {
                Write-Warning "Keine Domänencontroller in der Domäne $DomainName gefunden."
            }
        } catch {
            Write-Warning "Fehler beim Abrufen der Domänencontroller für Domäne $DomainName: $($_.Exception.Message)"
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
            # repadmin /syncall <DCName> <Optionen>
            # /A : Synchronisiert von allen Replikationspartnern (All).
            # /e : Über Standortgrenzen hinweg (Enterprise). Dies ist wichtig für den gesamten Forest.
            # /d : Identifiziert Server anhand ihres Distinguished Name in Nachrichten (detaillierter).
            # /P : (Optional) Pusht Änderungen vom angegebenen DC zu seinen Partnern.
            #      Ohne /P initiiert der DC einen Pull von seinen Partnern. Für das Ziel,
            #      dass alle DCs aktuelle Daten haben, ist ein Pull auf jedem DC (ohne /P hier)
            #      oder ein Push (/P) von allen DCs eine valide Strategie.
            #      Ein Pull auf jedem DC ist oft der Standardansatz, um sicherzustellen,
            #      dass jeder DC die neuesten Informationen abruft.
            
            # Diese Zeile weist den DC an, Änderungen von all seinen Partnern zu ziehen (pull).
            $Command = "repadmin /syncall $DCHostname /A /e /d"
            Write-Host "Führe aus: $Command"
            Invoke-Expression $Command

            # Kurze Pause, um die Ausgabe abzuwarten oder bei sehr vielen DCs die Last etwas zu verteilen.
            # Start-Sleep -Seconds 1 
            
            # Überprüfung (optional, da repadmin selbst Fehler ausgibt)
            # Man könnte `repadmin /showrepl $DCHostname` analysieren, aber das wird schnell komplex.
            # Für dieses Skript verlassen wir uns auf die Ausgabe von repadmin /syncall.
            # Eine erfolgreich angestoßene Replikation bedeutet nicht zwingend, dass sofort alle Daten da sind,
            # sondern dass der Prozess initiiert wurde.

            Write-Host "Replikation für $DCHostname erfolgreich angestoßen." -ForegroundColor Green
            $SuccessCount++
        } catch {
            Write-Warning "Fehler beim Ausführen von repadmin für $DCHostname: $($_.Exception.Message)"
            $FailCount++
            $FailedDCs += $DCHostname
        }
    }

    Write-Host "--------------------------------------------------"
    Write-Host "`nZusammenfassung des Replikationsanstoßes:"
    Write-Host "Erfolgreich angestoßen für: $SuccessCount Domänencontroller." -ForegroundColor Green
    if ($FailCount -gt 0) {
        Write-Warning "Fehlgeschlagen für: $FailCount Domänencontroller."
        Write-Warning "Fehlgeschlagene DCs: $($FailedDCs -join ', ')"
        Write-Warning "Überprüfen Sie die Fehlermeldungen oben und stellen Sie sicher, dass 'repadmin.exe' im Pfad ist und Netzwerkverbindung sowie Berechtigungen vorhanden sind."
    }
    Write-Host "Der Replikationsprozess im Hintergrund kann je nach Netzwerk und Auslastung noch einige Zeit in Anspruch nehmen."

} catch {
    Write-Error "Ein schwerwiegender Fehler ist aufgetreten: $($_.Exception.Message)"
    Write-Error "Stellen Sie sicher, dass das ActiveDirectory Modul installiert ist (Teil der Remote Server Administration Tools - RSAT) und Sie über ausreichende Berechtigungen verfügen (z.B. Mitglied der Gruppe 'Domänen-Admins' oder 'Organisations-Admins')."
}
