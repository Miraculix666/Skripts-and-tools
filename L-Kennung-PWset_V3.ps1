function Protokollieren {
    param (
        [string]$Nachricht,
        [ValidateSet("INFO","WARNUNG","FEHLER")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss.fff"
    $logEntry = "$timestamp [$Level] - $Nachricht"
    
    # Create directory structure if missing
    $logDirectory = Split-Path $LogPath -Parent
    if (-not (Test-Path -Path $logDirectory)) {
        New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
    }
    
    # Write to log file
    Add-Content -Path $LogPath -Value $logEntry -Encoding UTF8
    
    # Console output
    if (-not $Silent) {
        Write-Host $logEntry -ForegroundColor @{"INFO"="White";"WARNUNG"="Yellow";"FEHLER"="Red"}[$Level]
    }
}
