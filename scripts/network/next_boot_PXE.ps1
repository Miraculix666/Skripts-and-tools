# Erhalte die BIOS-Einstellungen
$BiosSettings = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class HpBIOSSetting

# Suche nach der Boot-Reihenfolge
$BootOrderSetting = $BiosSettings | Where-Object { $_.Name -eq 'Boot Order' }

# Zeige die Boot-Reihenfolge an
$BootOrderSetting.StringValue


# Setze die Boot-Reihenfolge auf PXE zuerst
bcdedit /set {bootmgr} bootsequence PXE

# Mache die Änderung dauerhaft
bcdedit /set {bootmgr} persistent yes


# Setze die Boot-Reihenfolge auf PXE zuerst
bcdedit /set {bootmgr} bootsequence PXE

# Mache die Änderung dauerhaft
bcdedit /set {bootmgr} persistent yes

# Erhalte die BIOS-Einstellungen
$BiosSettings = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class HpBIOSSetting

# Suche nach der Boot-Reihenfolge
$BootOrderSetting = $BiosSettings | Where-Object { $_.Name -eq 'Boot Order' }

# Setze die Boot-Reihenfolge auf PXE zuerst
$BootOrderSetting.StringValue = 'PXE Boot, Hard Drive, CD-ROM, Floppy'
$BootOrderSetting.Put()



