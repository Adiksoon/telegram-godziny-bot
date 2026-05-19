$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

if (!(Test-Path ".\.env")) {
    Copy-Item ".\.env.example" ".\.env"
    Write-Host "Utworzono .env. Wklej token bota do pliku .env i uruchom ponownie."
    exit 1
}

.\.venv\Scripts\python.exe .\bot.py
