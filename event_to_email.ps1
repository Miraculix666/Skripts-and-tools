# Erstellen Sie ein neues Outlook-Anwendungsobjekt
$outlook = New-Object -ComObject Outlook.Application

# Erstellen Sie einen neuen Termin
$appointment = $outlook.CreateItem(1) # 1 = olAppointmentItem

# Legen Sie die Details des Termins fest
$appointment.Subject = "Betreff des Termins"
$appointment.Body = "Beschreibung des Termins"
$appointment.Start = [datetime]::Now.AddDays(1)
$appointment.Duration = 60 # Dauer in Minuten

# Speichern Sie den Termin
$appointment.Save()

# Erstellen Sie eine neue E-Mail
$mail = $outlook.CreateItem(0) # 0 = olMailItem

# Legen Sie die Details der E-Mail fest
$mail.Subject = "Neuer Termin: " + $appointment.Subject
$mail.Body = "Hier sind die Details des neuen Termins:`n`n" +
             $appointment.Body + "`n`n" +
             "Startzeit: " + $appointment.Start.ToString() + "`n" +
             "Dauer: " + $appointment.Duration.ToString() + " Minuten"

# Fügen Sie den Termin als Anhang hinzu
$mail.Attachments.Add($appointment)

# Legen Sie den Empfänger der E-Mail fest
$mail.Recipients.Add("beispiel@domain.com") # Ändern Sie die Adresse in Ihre eigene

# Senden Sie die E-Mail
$mail.Send()
