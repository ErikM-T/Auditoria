$url = "https://github.com/ErikM-T/Auditoria/releases/latest/download/AuditoriaEquipos.exe"
$exe = "$env:TEMP\AuditoriaEquipos.exe"

iwr $url -o $exe
&$exe