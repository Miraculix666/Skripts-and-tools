$weird = ConvertTo-SecureString -AsPlainText "System.Security.SecureString" -Force
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($weird)
$PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
Write-Host "Weird password is: $PlainPassword"
