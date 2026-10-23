$url = "https://github.com/ErikM-T/Auditoria/releases/latest/download/AuditoriaEquipos.exe"
$exe = "$env:TEMP\AuditoriaEquipos.exe"

Invoke-WebRequest -Uri $url -OutFile $exe -UseBasicParsing

Unblock-File -Path $exe -ErrorAction SilentlyContinue

Start-Process -FilePath $exe -Wait

Start-Job -ScriptBlock {
    param($path)
    Start-Sleep -Seconds 2
    if (Test-Path $path) {
        Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
    }
} -ArgumentList $exe | Out-Null