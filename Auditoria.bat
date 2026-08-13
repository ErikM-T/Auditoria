<# :
@echo off
:: Requerir permisos de Administrador de forma automática
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Solicitando privilegios de Administrador...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Cambiar al directorio donde está guardado el archivo .bat
cd /d "%~dp0"

setlocal
title Inspeccion de Seguridad y Auditoria
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression ([System.IO.File]::ReadAllText('%~f0'))"
exit /b
#>

Add-Type -AssemblyName PresentationFramework, System.Windows.Forms, System.Drawing, Microsoft.VisualBasic

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Inspección de Seguridad y Auditoría de Equipos" Height="720" Width="800"
        WindowStartupLocation="CenterScreen" ResizeMode="CanMinimize" Background="#181818">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Encabezado -->
        <StackPanel Grid.Row="0" Margin="0,0,0,10">
            <TextBlock Text="INSPECCIÓN DE SEGURIDAD Y AUDITORÍA DE TI" Foreground="#00E5FF" FontSize="18" FontWeight="Bold" HorizontalAlignment="Center"/>
            <TextBlock Text="Evaluación automática y verificación técnico-operativa" Foreground="#AAAAAA" FontSize="12" HorizontalAlignment="Center" Margin="0,2,0,0"/>
        </StackPanel>

        <!-- Área de Texto de Resultados -->
        <TextBox Name="TxtResultados" Grid.Row="1" AcceptsReturn="True" TextWrapping="Wrap" IsReadOnly="True" Background="#1E1E1E" Foreground="#00FF66" FontFamily="Consolas" FontSize="12" Padding="10" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>

        <!-- Panel Inferior de Botones -->
        <Grid Grid.Row="2" Margin="0,10,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button Name="BtnExportarExcel" Grid.Column="0" Content="Exportar Formato Excel GE-FO-21" Height="38" Background="#16A34A" Foreground="White" FontWeight="Bold" Cursor="Hand"/>
            <Button Name="BtnCerrar" Grid.Column="2" Content="Cerrar Ventana" Height="38" Background="#DC2626" Foreground="White" FontWeight="Bold" Cursor="Hand"/>
        </Grid>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$TxtResultados    = $window.FindName("TxtResultados")
$BtnExportarExcel = $window.FindName("BtnExportarExcel")
$BtnCerrar        = $window.FindName("BtnCerrar")

$global:AuditData = [ordered]@{}

function Do-Events {
    [System.Windows.Forms.Application]::DoEvents()
}

function Test-WebSiteAccess {
    param([string]$Url, [string]$Domain = "")
    try {
        $response = Invoke-WebRequest -Uri $Url -TimeoutSec 4 -UseBasicParsing -ErrorAction Stop
        if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 400) { return $true }
        return $false
    } catch {
        if ($Domain) {
            $tcpTest = Test-NetConnection -ComputerName $Domain -Port 443 -WarningAction SilentlyContinue
            if ($tcpTest.TcpTestSucceeded) { return $true }
        }
        return $false
    }
}

# LÓGICA PRINCIPAL DE AUDITORÍA
$EjecutarAuditoria = {
    $TxtResultados.Text = "Iniciando inspección de seguridad en modo Administrador...`r`n`r`n"
    Do-Events

    # 1. DATOS ADMINISTRATIVOS
    $TxtResultados.Text += "[Paso 1] Registrando información administrativa...`r`n"
    Do-Events
    $proceso     = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese Nombre del Proceso o Área:", "1. Proceso / Área", "DESARROLLO / SOPORTE")
    $ejecutado   = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese Nombre de quien realiza la inspección:", "2. Ejecutado por", "ERIK ALONSO MACIAS HOYOS")
    $cargo       = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese Cargo de quien realiza la inspección:", "3. Cargo", "¿?")
    $responsable = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese Nombre del funcionario responsable del equipo:", "4. Responsable del Equipo", "")

    # Convertir a mayúsculas
    if ($proceso) { $proceso = $proceso.ToUpper() }
    if ($ejecutado) { $ejecutado = $ejecutado.ToUpper() }
    if ($cargo) { $cargo = $cargo.ToUpper() }
    if ($responsable) { $responsable = $responsable.ToUpper() }

    # 2. INFORMACIÓN BÁSICA DEL SISTEMA
    $TxtResultados.Text += "[Paso 2] Obteniendo datos del sistema y hardware...`r`n"
    Do-Events
    $sys = Get-CimInstance Win32_ComputerSystem
    $os  = Get-CimInstance Win32_OperatingSystem
    $userDom = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $tipoEq = if ($sys.PCSystemType -eq 2) { "Portátil" } else { "Escritorio" }

    # 3. PRIVILEGIOS DE ADMINISTRADOR
    $TxtResultados.Text += "[Paso 3] Auditando privilegios de cuenta de usuario...`r`n"
    Do-Events
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    $isAdminRole = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    $adminSid = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-544")
    $isInAdminGroup = $identity.Groups -contains $adminSid

    $strAdmin = if ($isAdminRole -or $isInAdminGroup) { "SI" } else { "NO" }

    # 4. RESTRICCIONES DE PERIFÉRICOS
    $TxtResultados.Text += "[Paso 4] Validando restricciones de periféricos...`r`n"
    Do-Events

    $usbReg = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR" -Name "Start" -ErrorAction SilentlyContinue).Start
    $usbDetec = if ($usbReg -eq 4) { "Bloqueado" } else { "Habilitado" }

    $msgUsb = "¿Los puertos USB de almacenamiento se encuentran HABILITADOS en este equipo?`r`n`r`n" +
              "• Detección del sistema (Registro): $usbDetec`r`n`r`n" +
              "• [Sí]: Registrar como 'Habilitado / No Restringido'`r`n" +
              "• [No]: Registrar como 'Bloqueado Correctamente'"

    $resUsbConfirm = [System.Windows.Forms.MessageBox]::Show($msgUsb, "Confirmación de Estado USB", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    $usbState = if ($resUsbConfirm -eq [System.Windows.Forms.DialogResult]::Yes) { "SI" } else { "NO" }

    $cdReg = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\cdrom" -Name "Start" -ErrorAction SilentlyContinue).Start
    $cdState = if ($cdReg -eq 4) { "NO" } else { "SI" }

    $btServ = Get-Service -Name "bthserv" -ErrorAction SilentlyContinue
    $btState = if ($btServ -and $btServ.Status -eq "Running") { "SI" } else { "NO" }

    $wifiAdapters = Get-NetAdapter -Physical | Where-Object { $_.MediaType -match "Native 802.11|Wireless" -and $_.Status -eq "Up" }
    $wifiState = if ($wifiAdapters) { "SI" } else { "NO" }

    # 5. VALIDACIÓN NAVEGADORES, SERVICIOS DE GOOGLE Y SITIOS WEB
    $TxtResultados.Text += "[Paso 5] Evaluando políticas de navegación, servicios Google y accesos web...`r`n"
    Do-Events

    # Pregunta sobre si el equipo cuenta con libre acceso a navegación web
    $msgNavLibre = "¿El equipo cuenta con acceso LIBRE a navegación web (sin restricciones de políticas o filtros)?"
    $resNavLibre = [System.Windows.Forms.MessageBox]::Show($msgNavLibre, "Navegación Web Libre", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    $navState = if ($resNavLibre -eq [System.Windows.Forms.DialogResult]::Yes) { "SI" } else { "NO" }

    # Pruebas de acceso a servicios web de correo y Google
    $outlookOk = Test-WebSiteAccess -Url "https://outlook.live.com" -Domain "outlook.live.com"
    $gmailOk   = Test-WebSiteAccess -Url "https://mail.google.com" -Domain "mail.google.com"
    $meetOk    = Test-WebSiteAccess -Url "https://meet.google.com" -Domain "meet.google.com"
    $driveOk   = Test-WebSiteAccess -Url "https://drive.google.com" -Domain "drive.google.com"
    $youtubeOk = Test-WebSiteAccess -Url "https://youtube.com"       -Domain "youtube.com"

    $outlookState = if ($outlookOk) { "SI" } else { "NO" }

    # Compilar lista de servicios de Google y correo activos
    $serviciosMailList = @()
    if ($gmailOk)   { $serviciosMailList += "Gmail" }
    if ($meetOk)    { $serviciosMailList += "Meet" }
    if ($driveOk)   { $serviciosMailList += "Drive" }
    if ($youtubeOk) { $serviciosMailList += "YouTube" }
    if ($outlookOk) { $serviciosMailList += "Outlook" }

    $gmailState = if ($gmailOk -or $meetOk -or $driveOk -or $youtubeOk) { "SI" } else { "NO" }
    $serviciosMailStr = if ($serviciosMailList.Count -gt 0) { ($serviciosMailList -join ", ") } else { "Ninguno disponible" }

    # Pruebas y consulta de Almacenamiento en la nube
    $cloudSites = @(
        @{ Url = "https://drive.google.com"; Domain = "drive.google.com" },
        @{ Url = "https://dropbox.com";      Domain = "dropbox.com" },
        @{ Url = "https://wetransfer.com";   Domain = "wetransfer.com" },
        @{ Url = "https://mega.io";          Domain = "mega.io" }
    )
    $blockedCount = 0
    foreach ($site in $cloudSites) {
        if (-not (Test-WebSiteAccess -Url $site.Url -Domain $site.Domain)) { $blockedCount++ }
    }
    $cloudState = if ($blockedCount -lt 2) { "SI" } else { "NO" }

    # Cuadro para ingresar el servicio de almacenamiento que está utilizando
    $cloudServiceUsed = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese el servicio de almacenamiento en la nube que está usando en el equipo (Ej: Google Drive, OneDrive, Mega, N/A):", "Servicios de Almacenamiento", "Google Drive")
    if ([string]::IsNullOrWhiteSpace($cloudServiceUsed)) { $cloudServiceUsed = "N/A" } else { $cloudServiceUsed = $cloudServiceUsed.ToUpper() }

    $resHisto = [System.Windows.Forms.MessageBox]::Show("¿El historial de navegación contiene sitios NO autorizados o de uso personal?", "Historial Web", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    $histoState = if ($resHisto -eq [System.Windows.Forms.DialogResult]::Yes) { "SI" } else { "NO" }

    # 6. ANTIVIRUS, VPN Y RDP
    $TxtResultados.Text += "[Paso 6] Auditando Antivirus, VPN y RDP...`r`n"
    Do-Events

    $av = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct -ErrorAction SilentlyContinue
    $avName = if ($av) { ($av.displayName -join ", ").ToUpper() } else { "WINDOWS DEFENDER" }

    $avVersion = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese la versión del Antivirus instalado ($avName):", "Versión del Antivirus", "Última Versión / N/A")
    if (-not $avVersion) { $avVersion = "N/A" } else { $avVersion = $avVersion.ToUpper() }

    $nowStr = Get-Date -Format "dd/MM/yyyy"
    $avLastUpdate = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese la Fecha de última actualización de firmas (DD/MM/YYYY):", "Fecha Actualización Antivirus", $nowStr)
    if (-not $avLastUpdate) { $avLastUpdate = "No registrada" } else { $avLastUpdate = $avLastUpdate.ToUpper() }

    $vpn = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match "GlobalProtect|Fortinet|TAP|VPN|Cisco" -or $_.Name -match "VPN" }
    $rdpReg = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -ErrorAction SilentlyContinue).fDenyTSConnections
    $vpnRdpState = if ($vpn -or $rdpReg -eq 0) { "SI" } else { "NO" }

    # 7. ARCHIVOS Y MULTIMEDIA
    $TxtResultados.Text += "[Paso 7] Inspeccionando archivos y Papelera...`r`n"
    Do-Events
    $deskPath = [Environment]::GetFolderPath("Desktop")
    $deskFiles = (Get-ChildItem -Path $deskPath -File | Where-Object { $_.Extension -notmatch "\.lnk$|\.url$" }).Count

    $shell = New-Object -ComObject Shell.Application
    $recycleBin = $shell.NameSpace(0xa)
    $recycleCount = $recycleBin.Items().Count

    $userProfile = $env:USERPROFILE
    $picsCount  = (Get-ChildItem -Path "$userProfile\Pictures" -File -Recurse -ErrorAction SilentlyContinue).Count
    $musicCount = (Get-ChildItem -Path "$userProfile\Music" -File -Recurse -ErrorAction SilentlyContinue).Count
    $videoCount = (Get-ChildItem -Path "$userProfile\Videos" -File -Recurse -ErrorAction SilentlyContinue).Count
    $docsCount  = (Get-ChildItem -Path "$userProfile\Documents" -File -Recurse -ErrorAction SilentlyContinue).Count

    # 8. CIFRADO DISCO Y NTP
    $TxtResultados.Text += "[Paso 8] Auditando BitLocker y Sincronización NTP...`r`n"
    Do-Events
    $bitlocker = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
    $cifradoState = if ($bitlocker -and $bitlocker.ProtectionStatus -eq "On") { "SI" } else { "NO" }

    # Comprobación de Sincronización NTP
    $ntpOut = w32tm /query /status 2>&1
    $ntpSource = ($ntpOut | Select-String "Origen:" -SimpleMatch) -replace "Origen:", ""
    $ntpState = if ($ntpSource -and $ntpSource -notmatch "Local RAM|Free-running Local Clock") { "SI" } else { "NO" }

    # 9. SOFTWARE NO AUTORIZADO, AIP Y CONSTRASEÑAS
    $TxtResultados.Text += "[Paso 9] Evaluando AIP y Contraseñas Físicas...`r`n"
    Do-Events
    $resApps = [System.Windows.Forms.MessageBox]::Show("¿Se detectaron aplicaciones NO autorizadas instaladas?", "Software No Autorizado", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    $appsState = if ($resApps -eq [System.Windows.Forms.DialogResult]::Yes) { "SI" } else { "NO" }

    $resAIP  = [System.Windows.Forms.MessageBox]::Show("¿Azure Information Protect (AIP) está habilitado?", "Seguridad AIP", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    $aipState = if ($resAIP -eq [System.Windows.Forms.DialogResult]::Yes) { "SI" } else { "NO" }

    $resPass = [System.Windows.Forms.MessageBox]::Show("¿Se identificaron contraseñas visibles o escritas (post-it, notas, etc.)?", "Seguridad Física", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    $passState = if ($resPass -eq [System.Windows.Forms.DialogResult]::Yes) { "SI" } else { "NO" }

    # CONSOLIDACIÓN DE DATOS
    $global:AuditData = [ordered]@{
        "Proceso"          = $proceso
        "EjecutadoPor"      = $ejecutado
        "Cargo"            = $cargo
        "Responsable"      = $responsable
        "UsuarioDominio"   = $userDom.ToUpper()
        "NombreEquipo"     = $env:COMPUTERNAME.ToUpper()
        "VersionSO"        = "$($os.Caption) (Build $($os.BuildNumber))".ToUpper()
        "TipoEquipo"       = $tipoEq
        "IsAdmin"          = $strAdmin
        "USB"              = $usbState
        "CD"               = $cdState
        "Bluetooth"        = $btState
        "WIFI"             = $wifiState
        "NavWeb"           = $navState
        "Outlook"          = $outlookState
        "Gmail"            = $gmailState
        "GmailServices"    = $serviciosMailStr.ToUpper()
        "Cloud"            = $cloudState
        "CloudServiceUsed" = $cloudServiceUsed
        "Historial"        = $histoState
        "AVName"           = $avName
        "AVVer"            = $avVersion
        "AVUpdate"         = $avLastUpdate
        "AppsUnauth"       = $appsState
        "VpnRdp"           = $vpnRdpState
        "DeskFiles"        = $deskFiles
        "RecycleFiles"     = $recycleCount
        "PicsCount"        = $picsCount
        "MusicCount"       = $musicCount
        "VideoCount"       = $videoCount
        "DocsCount"        = $docsCount
        "BitLocker"        = $cifradoState
        "NTPState"         = $ntpState
        "AIP"              = $aipState
        "PassVisible"      = $passState
    }

    $TxtResultados.Text += "`r`n=== INSPECCIÓN FINALIZADA. LISTO PARA EXPORTAR A EXCEL ==="
}

$window.Add_ContentRendered({
    &$EjecutarAuditoria
})

# EXPORTACIÓN
$BtnExportarExcel.Add_Click({
    if ($global:AuditData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No hay datos capturados.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Asegurar módulo ImportExcel (Descarga automática aceptando licencias)
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false -ErrorAction SilentlyContinue
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -AcceptLicense -Confirm:$false -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No se pudo instalar el paquete de lectura/escritura de Excel. Asegúrate de tener conexión a Internet.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }
    Import-Module ImportExcel

    $nombrePersona = [string]$global:AuditData["Responsable"]
    if ([string]::IsNullOrWhiteSpace($nombrePersona)) { $nombrePersona = "Usuario" }

    # BÚSQUEDA Y FALLBACK DE PLANTILLA DESDE GOOGLE DRIVE
    $dirBase = (Get-Location).Path
    $nombrePlantillaExacto = "GE-FO-21 INSPECCION DE SEGURIDAD DE EQUIPOS (nombre).xlsx"
    $rutaPlantilla = Join-Path $dirBase $nombrePlantillaExacto

    if (-not (Test-Path -Path $rutaPlantilla)) {
        $posiblePlantilla = Get-ChildItem -Path $dirBase -Filter "GE-FO-21*.xlsx" | Select-Object -First 1
        if ($posiblePlantilla) { $rutaPlantilla = $posiblePlantilla.FullName }
    }

    # Si NO existe la plantilla en la carpeta local, se descarga automáticamente desde Google Drive
    if (-not (Test-Path -Path $rutaPlantilla)) {
        try {
            $TxtResultados.Text += "`r`n[Drive] Plantilla local no encontrada. Descargando desde Google Drive...`r`n"
            Do-Events

            $driveFileId = "1O08J0U7lHBwVbQidkZlQC-pN6yC96vWs"
            $driveDownloadUrl = "https://drive.google.com/uc?export=download&id=$driveFileId"

            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $driveDownloadUrl -OutFile $rutaPlantilla -UserAgent "Mozilla/5.0" -ErrorAction Stop

            $TxtResultados.Text += "[Drive] Plantilla descargada exitosamente desde la nube.`r`n"
            Do-Events
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No se encontró la plantilla localmente y falló la descarga desde Google Drive: $_", "Error de Plantilla", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }

    # Nombre del archivo final SIN paréntesis
    $nombreLimpio = ($nombrePersona -replace '[\\/:*?"<>|]', '').Trim()
    $nombreArchivoFinal = "GE-FO-21 INSPECCION DE SEGURIDAD DE EQUIPOS $nombreLimpio.xlsx"
    $rutaSalida = Join-Path (Split-Path $rutaPlantilla -Parent) $nombreArchivoFinal

    try {
        Copy-Item -Path $rutaPlantilla -Destination $rutaSalida -Force

        # Edición interna de Excel
        $excelPkg = Open-ExcelPackage -Path $rutaSalida
        $ws = $excelPkg.Workbook.Worksheets[1]

        function Set-CheckMarkDirect ($row, $value) {
            $valStr = [string]$value
            if ($valStr -eq "SI") {
                $ws.Cells["C$row"].Value = "X"
            } elseif ($valStr -eq "NO") {
                $ws.Cells["D$row"].Value = "X"
            } else {
                $ws.Cells["E$row"].Value = "X"
            }
        }

        function Set-FileCheckMarkDirect ($row, $count) {
            if ([int]$count -gt 0) {
                $ws.Cells["C$row"].Value = "X"
            } else {
                $ws.Cells["D$row"].Value = "X"
            }
        }

        # Encabezado
        $ws.Cells["B7"].Value = [string]($global:AuditData["Proceso"] -join " ")
        $ws.Cells["G7"].Value = (Get-Date -Format "dd/MM/yyyy")
        $ws.Cells["B8"].Value = [string]($global:AuditData["EjecutadoPor"] -join " ")
        $ws.Cells["G8"].Value = [string]($global:AuditData["Cargo"] -join " ")

        # Detalle de Equipo
        $ws.Cells["B11"].Value = [string]($global:AuditData["Responsable"] -join " ")
        $ws.Cells["B12"].Value = [string]($global:AuditData["UsuarioDominio"] -join " ")
        $ws.Cells["B13"].Value = [string]($global:AuditData["NombreEquipo"] -join " ")
        $ws.Cells["B14"].Value = [string]($global:AuditData["VersionSO"] -join " ")

        # Tipo de Equipo
        if ([string]$global:AuditData["TipoEquipo"] -eq "Portátil") {
            $ws.Cells["C15"].Value = "X (Portátil)"
        } else {
            $ws.Cells["E15"].Value = "X (Escritorio)"
        }

        # Verificaciones
        Set-CheckMarkDirect 19 $global:AuditData["IsAdmin"]
        Set-CheckMarkDirect 20 $global:AuditData["USB"]
        Set-CheckMarkDirect 21 $global:AuditData["CD"]
        Set-CheckMarkDirect 22 $global:AuditData["Bluetooth"]
        Set-CheckMarkDirect 23 $global:AuditData["WIFI"]
        Set-CheckMarkDirect 24 $global:AuditData["NavWeb"]
        Set-CheckMarkDirect 25 $global:AuditData["Outlook"]

        # Gmail / Servicios de Correo y Google
        Set-CheckMarkDirect 26 $global:AuditData["Gmail"]
        $ws.Cells["F26"].Value = [string]($global:AuditData["GmailServices"] -join " ")

        # Almacenamiento en la Nube
        Set-CheckMarkDirect 27 $global:AuditData["Cloud"]
        $ws.Cells["F27"].Value = [string]($global:AuditData["CloudServiceUsed"] -join " ")

        Set-CheckMarkDirect 28 $global:AuditData["Historial"]

        # Antivirus
        $avNombreStr  = [string]($global:AuditData['AVName'] -join ", ")
        $avVersionStr = [string]($global:AuditData['AVVer'] -join ", ")
        $ws.Cells["C29"].Value = "X"
        $ws.Cells["F29"].Value = "$avNombreStr (Ver: $avVersionStr)"
        $ws.Cells["C30"].Value = [string]($global:AuditData["AVUpdate"] -join " ")

        Set-CheckMarkDirect 32 $global:AuditData["AppsUnauth"]
        Set-CheckMarkDirect 33 $global:AuditData["VpnRdp"]

        # Conteo de Archivos
        Set-FileCheckMarkDirect 34 $global:AuditData['DeskFiles']
        $ws.Cells["F34"].Value = "$([string]$global:AuditData['DeskFiles']) archivos"

        Set-FileCheckMarkDirect 35 $global:AuditData['RecycleFiles']
        $ws.Cells["F35"].Value = "$([string]$global:AuditData['RecycleFiles']) archivos"

        Set-FileCheckMarkDirect 36 $global:AuditData['PicsCount']
        $ws.Cells["F36"].Value = "$([string]$global:AuditData['PicsCount']) archivos"

        Set-FileCheckMarkDirect 37 $global:AuditData['MusicCount']
        $ws.Cells["F37"].Value = "$([string]$global:AuditData['MusicCount']) archivos"

        Set-FileCheckMarkDirect 38 $global:AuditData['VideoCount']
        $ws.Cells["F38"].Value = "$([string]$global:AuditData['VideoCount']) archivos"

        Set-FileCheckMarkDirect 39 $global:AuditData['DocsCount']
        $ws.Cells["F39"].Value = "$([string]$global:AuditData['DocsCount']) archivos"

        # BitLocker
        Set-CheckMarkDirect 40 $global:AuditData["BitLocker"]

        # Sincronización NTP
        Set-CheckMarkDirect 41 $global:AuditData["NTPState"]
        $ws.Cells["F41"].Value = "NTP"

        # AIP y Contraseñas
        Set-CheckMarkDirect 42 $global:AuditData["AIP"]
        Set-CheckMarkDirect 43 $global:AuditData["AIP"]
        Set-CheckMarkDirect 44 $global:AuditData["AIP"]
        Set-CheckMarkDirect 45 $global:AuditData["PassVisible"]

        # Guardar y cerrar paquete
        Close-ExcelPackage $excelPkg

        # SUBIDA AUTOMÁTICA DEL EXCEL DILIGENCIADO A GOOGLE DRIVE
        $scriptUrl = "https://script.google.com/macros/s/AKfycbx1UgHhmeBbWUJrfigweQ3xpMv0ZRaTWWxW0GQDwVBkdnDxtLLcXQK3YlpGdQSAcDdF/exec"

        $driveStatus = ""
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            
            $fileBytes = [System.IO.File]::ReadAllBytes($rutaSalida)
            $base64File = [System.Convert]::ToBase64String($fileBytes)

            $jsonBody = @{
                fileName    = $nombreArchivoFinal
                mimeType    = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                fileContent = $base64File
            } | ConvertTo-Json -Depth 3

            $response = Invoke-RestMethod -Uri $scriptUrl -Method Post -Body $jsonBody -ContentType "application/json"

            if ($response.status -eq "success") {
                $driveStatus = "`r`n`r`n¡Reporte SUBIDO EXITOSAMENTE a Google Drive!"
            } else {
                $driveStatus = "`r`n`r`nNo se pudo subir a Google Drive: $($response.message)"
            }
        } catch {
            $driveStatus = "`r`n`r`nError de conexión al subir a Google Drive: $_"
        }

        [System.Windows.Forms.MessageBox]::Show("Reporte generado exitosamente con el nombre:`r`n$nombreArchivoFinal$driveStatus", "Exportación Completa", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error guardando el archivo Excel: $_", "Error de Exportación", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$BtnCerrar.Add_Click({ $window.Close() })
$null = $window.ShowDialog()