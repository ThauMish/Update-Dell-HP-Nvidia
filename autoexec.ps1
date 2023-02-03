Set-ExecutionPolicy Bypass -scope Process -Force
$scriptpath = "drivers.ps1"
Start-process PowerShell.exe -ArgumentList "-File $scriptpath"