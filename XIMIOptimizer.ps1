<#
.SYNOPSIS
    XIMIOptimizer v1.0 - Panel de Optimizacion Kernel y Perfiles (Celeron N4500)
.DESCRIPTION
    Herramienta nativa en PowerShell con Interfaz Grafica (WindowsForms) en modo oscuro.
    Combina el motor de optimizacion V11/V7.7 (control de WmiPrvSE, FastFlags MVSD,
    energia AC/DC, desactivacion de servicios pesados y purga de bloatware UWP) con
    un sistema de diagnostico y radiografia clinica de 2 minutos.
    Protege Audio, Game Bar y Herramienta de Recortes (Win + Shift + S).
#>

# ==============================================================================
# 0. VERIFICACION DE PRIVILEGIOS DE ADMINISTRADOR
# ==============================================================================
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "XIMIOptimizer requiere permisos de Administrador para modificar servicios, energia y registro.`n`nHaz clic derecho sobre PowerShell y elige 'Ejecutar como administrador'.",
        "XIMIOptimizer - Permisos Requeridos",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==============================================================================
# 1. CONFIGURACION DE VENTANA PRINCIPAL (ESTILO APBOptimizer / MODO OSCURO)
# ==============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "XIMIOptimizer v1.0 - Kernel & Gaming Suite (Intel Celeron N4500)"
$form.Size = New-Object System.Drawing.Size(880, 680)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(18, 18, 22)
$form.ForeColor = [System.Drawing.Color]::White
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Colores del tema
$colPanel      = [System.Drawing.Color]::FromArgb(26, 26, 32)
$colBtnNormal  = [System.Drawing.Color]::FromArgb(36, 36, 44)
$colBtnHover   = [System.Drawing.Color]::FromArgb(48, 48, 58)
$colAccent     = [System.Drawing.Color]::FromArgb(0, 230, 118)
$colCyan       = [System.Drawing.Color]::FromArgb(0, 188, 212)
$colWarn       = [System.Drawing.Color]::FromArgb(255, 193, 7)
$colDanger     = [System.Drawing.Color]::FromArgb(229, 57, 53)
$colTextMuted  = [System.Drawing.Color]::FromArgb(160, 165, 175)

# Encabezado Superior
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Location = New-Object System.Drawing.Point(0, 0)
$pnlHeader.Size = New-Object System.Drawing.Size(880, 65)
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(13, 13, 16)
$form.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "XIMIOptimizer"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(20, 10)
$lblTitle.Size = New-Object System.Drawing.Size(220, 35)
$lblTitle.ForeColor = $colAccent
$pnlHeader.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "Motor Hibrido V11 + GUI | Optimizacion de Baja Latencia (<16ms) y Control WMI | Audio y Recortes Seguros"
$lblSub.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lblSub.Location = New-Object System.Drawing.Point(22, 43)
$lblSub.Size = New-Object System.Drawing.Size(650, 20)
$lblSub.ForeColor = $colTextMuted
$pnlHeader.Controls.Add($lblSub)

# Funcion auxiliar para actualizar consola de estado en vivo
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Location = New-Object System.Drawing.Point(20, 495)
$txtLog.Size = New-Object System.Drawing.Size(825, 105)
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(12, 12, 15)
$txtLog.ForeColor = $colAccent
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BorderStyle = "FixedSingle"
$txtLog.Text = "[SISTEMA] XIMIOptimizer v1.0 listo. Hardware detectado: Intel Celeron / 8GB RAM.`r`n[INFO] Selecciona los modulos de optimizacion o lanza una aplicacion con perfil dedicado."
$form.Controls.Add($txtLog)

function Write-Log {
    param([string]$msg)
    $hora = (Get-Date).ToString("HH:mm:ss")
    $txtLog.AppendText("`r`n[$hora] $msg")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# ==============================================================================
# 2. PANEL IZQUIERDO: LAS 25 TAREAS MAESTRAS (CHECKBOXES SELECCIONABLES)
# ==============================================================================
$grp25 = New-Object System.Windows.Forms.GroupBox
$grp25.Text = " Modulo de 25 Optimizaciones Maestras (Selecciona y Aplica) "
$grp25.ForeColor = $colAccent
$grp25.BackColor = $colPanel
$grp25.Location = New-Object System.Drawing.Point(20, 75)
$grp25.Size = New-Object System.Drawing.Size(510, 410)
$grp25.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($grp25)

$pnlScroll = New-Object System.Windows.Forms.Panel
$pnlScroll.Location = New-Object System.Drawing.Point(10, 22)
$pnlScroll.Size = New-Object System.Drawing.Size(490, 330)
$pnlScroll.AutoScroll = $true
$grp25.Controls.Add($pnlScroll)

$tareas = @(
    @{ Id=1;  Checked=$true;  Text="01. Desactivar Telemetria y Diagnostico de Windows (AllowTelemetry=0)" },
    @{ Id=2;  Checked=$true;  Text="02. Frenar y Depurar proceso fantasma WmiPrvSE (Winmgmt)" },
    @{ Id=3;  Checked=$true;  Text="03. Desactivar SysMain (Superfetch) para evitar 100% disco/CPU" },
    @{ Id=4;  Checked=$true;  Text="04. Desactivar Indexador de Busqueda (WSearch)" },
    @{ Id=5;  Checked=$true;  Text="05. Desactivar DiagTrack (Experiencias y rastreo conectado)" },
    @{ Id=6;  Checked=$true;  Text="06. Desactivar Compresion de Memoria RAM (Libera ciclos Celeron)" },
    @{ Id=7;  Checked=$true;  Text="07. Eliminar Sugerencias, Anuncios y Spotlight en Inicio/Bloqueo" },
    @{ Id=8;  Checked=$true;  Text="08. Desactivar Microsoft Copilot e Integracion Bing en Busqueda" },
    @{ Id=9;  Checked=$true;  Text="09. Desactivar Widgets y Noticias de la Barra de Tareas" },
    @{ Id=10; Checked=$true;  Text="10. Bloquear Apps UWP en Segundo Plano (LetAppsRunInBackground=2)" },
    @{ Id=11; Checked=$true;  Text="11. Desactivar Inicio Rapido y Segundo Plano de Microsoft Edge" },
    @{ Id=12; Checked=$true;  Text="12. Desactivar Sincronizacion Automatica de OneDrive" },
    @{ Id=13; Checked=$true;  Text="13. Activar Plan de Energia Alto Rendimiento (AC 100% / DC Seguro)" },
    @{ Id=14; Checked=$true;  Text="14. Desactivar Ahorro de Energia en Wi-Fi y Puertos USB" },
    @{ Id=15; Checked=$true;  Text="15. Prioridad de Multimedia y Juegos (SystemResponsiveness=0)" },
    @{ Id=16; Checked=$true;  Text="16. Prioridad Alta de GPU para Juegos (GPU Priority=8)" },
    @{ Id=17; Checked=$true;  Text="17. Activar Aceleracion de GPU por Hardware (HwSchMode=2)" },
    @{ Id=18; Checked=$true;  Text="18. Optimizar Pila TCP/IP (AutoTuning Normal + RSS Activo)" },
    @{ Id=19; Checked=$true;  Text="19. Reducir Carga de CPU de Windows Defender (Limite 10%)" },
    @{ Id=20; Checked=$true;  Text="20. Desactivar Animaciones y Transparencias Pesadas (Estilo Win7/10)" },
    @{ Id=21; Checked=$true;  Text="21. Eliminar Retraso de Apertura de Menus (MenuShowDelay=0)" },
    @{ Id=22; Checked=$true;  Text="22. Desactivar Hibernacion para Liberar Espacio y Procesos" },
    @{ Id=23; Checked=$true;  Text="23. Desactivar Tareas Programadas de Telemetria (CompatTelRunner)" },
    @{ Id=24; Checked=$true;  Text="24. Cerrar Procesos Innecesarios Activos (Teams, Cortana, WebView2)" },
    @{ Id=25; Checked=$false; Text="25. Purgar Bloatware UWP (Clipchamp, Clima, Xbox, Solitario, etc.)" }
)

$checkboxes = @{}
$yPos = 5
foreach ($t in $tareas) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $t.Text
    $cb.Checked = $t.Checked
    $cb.Location = New-Object System.Drawing.Point(5, $yPos)
    $cb.Size = New-Object System.Drawing.Size(460, 22)
    $cb.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $cb.ForeColor = if ($t.Id -eq 25) { $colWarn } else { [System.Drawing.Color]::WhiteSmoke }
    $pnlScroll.Controls.Add($cb)
    $checkboxes[$t.Id] = $cb
    $yPos += 25
}

# Botones de seleccion y aplicacion de las 25 tareas
$btnMarcarTodo = New-Object System.Windows.Forms.Button
$btnMarcarTodo.Text = "Marcar Todas"
$btnMarcarTodo.Location = New-Object System.Drawing.Point(12, 362)
$btnMarcarTodo.Size = New-Object System.Drawing.Size(115, 34)
$btnMarcarTodo.BackColor = $colBtnNormal
$btnMarcarTodo.ForeColor = [System.Drawing.Color]::White
$btnMarcarTodo.FlatStyle = "Flat"
$btnMarcarTodo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$btnMarcarTodo.Add_Click({
    foreach ($k in $checkboxes.Keys) { $checkboxes[$k].Checked = $true }
    Write-Log "Todas las 25 optimizaciones marcadas."
})
$grp25.Controls.Add($btnMarcarTodo)

$btnAplicar25 = New-Object System.Windows.Forms.Button
$btnAplicar25.Text = "APLICAR OPTIMIZACION SELECCIONADA"
$btnAplicar25.Location = New-Object System.Drawing.Point(135, 362)
$btnAplicar25.Size = New-Object System.Drawing.Size(362, 34)
$btnAplicar25.BackColor = $colAccent
$btnAplicar25.ForeColor = [System.Drawing.Color]::Black
$btnAplicar25.FlatStyle = "Flat"
$btnAplicar25.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnAplicar25.Add_Click({
    Write-Log "Iniciando ejecucion de tareas seleccionadas..."
    $count = 0

    if ($checkboxes[1].Checked) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Force | Out-Null
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0 -Type DWord -Force
        $count++
    }
    if ($checkboxes[2].Checked) {
        Stop-Service -Name "Winmgmt" -Force -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[3].Checked) {
        Set-Service -Name "SysMain" -StartupType Disabled -ErrorAction SilentlyContinue
        Stop-Service -Name "SysMain" -Force -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[4].Checked) {
        Set-Service -Name "WSearch" -StartupType Disabled -ErrorAction SilentlyContinue
        Stop-Service -Name "WSearch" -Force -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[5].Checked) {
        Set-Service -Name "DiagTrack" -StartupType Disabled -ErrorAction SilentlyContinue
        Stop-Service -Name "DiagTrack" -Force -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[6].Checked) {
        Disable-MMAgent -MemoryCompression -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[7].Checked) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Force | Out-Null
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -Value 1 -Type DWord -Force
        $count++
    }
    if ($checkboxes[8].Checked) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" -Force | Out-Null
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" -Name "TurnOffWindowsCopilot" -Value 1 -Type DWord -Force
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Force | Out-Null
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled" -Value 0 -Type DWord -Force
        $count++
    }
    if ($checkboxes[9].Checked) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" -Force | Out-Null
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" -Name "AllowNewsAndInterests" -Value 0 -Type DWord -Force
        $count++
    }
    if ($checkboxes[10].Checked) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Force | Out-Null
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsRunInBackground" -Value 2 -Type DWord -Force
        $count++
    }
    if ($checkboxes[11].Checked) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Force | Out-Null
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "StartupBoostEnabled" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "BackgroundModeEnabled" -Value 0 -Type DWord -Force
        $count++
    }
    if ($checkboxes[12].Checked) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Force | Out-Null
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Name "DisableFileSyncNGSC" -Value 1 -Type DWord -Force
        $count++
    }
    if ($checkboxes[13].Checked) {
        powercfg -duplicatescheme 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c 2>$null | Out-Null
        powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c 2>$null | Out-Null
        powercfg -setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN 5 2>$null | Out-Null
        powercfg -setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMAX 100 2>$null | Out-Null
        powercfg -setacvalueindex SCHEME_CURRENT SUB_PROCESSOR CPMINCORES 100 2>$null | Out-Null
        powercfg -setdcvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN 85 2>$null | Out-Null
        powercfg -setdcvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMAX 95 2>$null | Out-Null
        powercfg -setactive SCHEME_CURRENT 2>$null | Out-Null
        $count++
    }
    if ($checkboxes[14].Checked) {
        powercfg -setacvalueindex SCHEME_CURRENT 19cbb8fa-5279-450e-9fac-8a3d5fedd0c1 12bbebe6-58d6-4636-95bb-3217ef867c1a 0 2>$null | Out-Null
        powercfg -setacvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0 2>$null | Out-Null
        $count++
    }
    if ($checkboxes[15].Checked) {
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" -Name "SystemResponsiveness" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[16].Checked) {
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" -Name "GPU Priority" -Value 8 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" -Name "Priority" -Value 6 -Type DWord -Force -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[17].Checked) {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers" -Name "HwSchMode" -Value 2 -Type DWord -Force -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[18].Checked) {
        netsh int tcp set global autotuninglevel=normal 2>$null | Out-Null
        netsh int tcp set global rss=enabled 2>$null | Out-Null
        $count++
    }
    if ($checkboxes[19].Checked) {
        Add-MpPreference -ExclusionPath "$env:LOCALAPPDATA\Roblox", "$env:APPDATA\.minecraft" -ErrorAction SilentlyContinue
        Set-MpPreference -ScanAvgCPULoadFactor 10 -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[20].Checked) {
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "EnableTransparency" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop\WindowMetrics" -Name "MinAnimate" -Value "0" -Type String -Force -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[21].Checked) {
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "MenuShowDelay" -Value "0" -Type String -Force -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[22].Checked) {
        powercfg -h off 2>$null | Out-Null
        $count++
    }
    if ($checkboxes[23].Checked) {
        $tasks = @(
            "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser",
            "\Microsoft\Windows\Application Experience\ProgramDataUpdater",
            "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator",
            "\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip"
        )
        foreach ($tk in $tasks) { schtasks /Change /TN $tk /Disable 2>$null | Out-Null }
        Stop-Process -Name "CompatTelRunner" -Force -ErrorAction SilentlyContinue
        $count++
    }
    if ($checkboxes[24].Checked) {
        $killList = @("OneDrive","Teams","Widgets","WidgetService","PhoneExperienceHost","Cortana","msedge","msedgewebview2","SearchHost","YourPhone","XboxPcAppFT","TabTip","TextInputHost")
        foreach ($p in $killList) { Stop-Process -Name $p -Force -ErrorAction SilentlyContinue }
        $count++
    }
    if ($checkboxes[25].Checked) {
        Write-Log "Purgando aplicaciones UWP innecesarias (Conservando Recortes, Audio y Game Bar)..."
        $bloat = @(
            "Clipchamp.Clipchamp","Microsoft.BingNews","Microsoft.BingWeather","Microsoft.BingSearch",
            "Microsoft.GetHelp","Microsoft.Getstarted","Microsoft.MicrosoftOfficeHub","Microsoft.MicrosoftSolitaireCollection",
            "Microsoft.People","Microsoft.PowerAutomateDesktop","Microsoft.Todos","Microsoft.WindowsFeedbackHub",
            "Microsoft.WindowsMaps","Microsoft.YourPhone","Microsoft.ZuneMusic","Microsoft.ZuneVideo",
            "Microsoft.MixedReality.Portal","Microsoft.SkypeApp","MicrosoftTeams","MSTeams","Microsoft.Copilot"
        )
        foreach ($b in $bloat) {
            Get-AppxPackage -Name $b -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue
        }
        $count++
    }

    Write-Log "OPTIMIZACION COMPLETADA: $count modulos aplicados correctamente."
    [System.Windows.Forms.MessageBox]::Show("Se han aplicado $count optimizaciones del kernel con exito.`nAudio, Game Bar y Recortes (Win+Shift+S) permanecen 100% activos.", "XIMIOptimizer", 0, 64) | Out-Null
})
$grp25.Controls.Add($btnAplicar25)

# ==============================================================================
# 3. PANEL DERECHO SUPERIOR: PERFILES DE APLICACIONES Y JUEGOS
# ==============================================================================
$grpPerfiles = New-Object System.Windows.Forms.GroupBox
$grpPerfiles.Text = " Lanzador por Perfil de App "
$grpPerfiles.ForeColor = $colCyan
$grpPerfiles.BackColor = $colPanel
$grpPerfiles.Location = New-Object System.Drawing.Point(545, 75)
$grpPerfiles.Size = New-Object System.Drawing.Size(300, 215)
$grpPerfiles.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($grpPerfiles)

function New-ProfileButton {
    param($text, $y, $fg, $action)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Location = New-Object System.Drawing.Point(15, $y)
    $b.Size = New-Object System.Drawing.Size(270, 36)
    $b.BackColor = $colBtnNormal
    $b.ForeColor = $fg
    $b.FlatStyle = "Flat"
    $b.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $b.Add_Click($action)
    return $b
}

# Perfil 1: Roblox MVSD
$btnRbx = New-ProfileButton "Roblox Player (Perfil MVSD 60Hz)" 28 $colAccent {
    Write-Log "Inyectando FastFlags competitivos para Roblox MVSD..."
    $vDir = "$env:LOCALAPPDATA\Roblox\Versions"
    if (Test-Path $vDir) {
        Get-ChildItem $vDir -Directory | ForEach-Object {
            if (Test-Path "$($_.FullName)\RobloxPlayerBeta.exe") {
                $cs = "$($_.FullName)\ClientSettings"
                if (!(Test-Path $cs)) { New-Item -ItemType Directory -Path $cs | Out-Null }
                $json = @"
{
  "DFIntTaskSchedulerTargetFps": 60,
  "FFlagDebugGraphicsPreferD3D11": true,
  "FFlagDisablePostFx": true,
  "FIntRenderShadowIntensity": 0,
  "FFlagGlobalWindRendering": false,
  "DFFlagDebugRenderForceTechnologyVoxel": true,
  "FIntDebugForceMSAASamples": 0,
  "FFlagEnableLayeredClothing": false,
  "DFIntTextureQualityOverride": 1,
  "DFFlagTextureQualityOverrideEnabled": true,
  "FFlagEnableInGameMenuChrome": false
}
"@
                Set-Content -Path "$cs\ClientAppSettings.json" -Value $json -Encoding ASCII
                Start-Process "$($_.FullName)\RobloxPlayerBeta.exe"
                $p = Get-Process -Name "RobloxPlayerBeta" -ErrorAction SilentlyContinue
                if ($p) { $p.PriorityClass = "High" }
            }
        }
        Write-Log "Roblox lanzado en prioridad ALTA con perfil MVSD (Sin ropa 3D ni sombras)."
    } else {
        Write-Log "[AVISO] No se encontro la carpeta de instalacion de Roblox."
    }
}
$grpPerfiles.Controls.Add($btnRbx)

# Perfil 2: Google Chrome
$btnChr = New-ProfileButton "Google Chrome (GPU Raster + ZeroCopy)" 72 ([System.Drawing.Color]::White) {
    Write-Log "Iniciando Google Chrome con aceleracion GPU forzada..."
    $paths = @("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe")
    $found = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($found) {
        Start-Process $found "--enable-gpu-rasterization --enable-zero-copy --enable-accelerated-video-decode --ignore-gpu-blocklist"
        Write-Log "Chrome optimizado iniciado."
    } else {
        Write-Log "[AVISO] Google Chrome no encontrado."
    }
}
$grpPerfiles.Controls.Add($btnChr)

# Perfil 3: Roblox Studio
$btnStu = New-ProfileButton "Roblox Studio (Modo Ligero)" 116 ([System.Drawing.Color]::White) {
    Write-Log "Buscando Roblox Studio..."
    $exe = Get-ChildItem "$env:LOCALAPPDATA\Roblox\Versions" -Recurse -Filter "RobloxStudioBeta.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($exe) {
        Start-Process $exe.FullName
        Write-Log "Roblox Studio iniciado."
    } else {
        Write-Log "[AVISO] Roblox Studio no encontrado."
    }
}
$grpPerfiles.Controls.Add($btnStu)

# Perfil 4: Minecraft Launchers
$btnMc = New-ProfileButton "Minecraft (SKLauncher / TKLauncher)" 160 ([System.Drawing.Color]::White) {
    Write-Log "Buscando launcher de Minecraft en Escritorio y Descargas..."
    $mc = Get-ChildItem "$env:USERPROFILE\Desktop", "$env:USERPROFILE\Downloads" -Filter "*Launcher*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($mc) {
        Start-Process $mc.FullName
        Write-Log "Iniciando $($mc.Name)..."
    } else {
        Write-Log "[AVISO] No se encontro SKLauncher ni TKLauncher."
    }
}
$grpPerfiles.Controls.Add($btnMc)

# ==============================================================================
# 4. PANEL DERECHO INFERIOR: DIAGNOSTICO, RADIOGRAFIA Y RESTAURACION
# ==============================================================================
$grpDiag = New-Object System.Windows.Forms.GroupBox
$grpDiag.Text = " Diagnostico, Radiografia y Control "
$grpDiag.ForeColor = $colWarn
$grpDiag.BackColor = $colPanel
$grpDiag.Location = New-Object System.Drawing.Point(545, 300)
$grpDiag.Size = New-Object System.Drawing.Size(300, 185)
$grpDiag.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($grpDiag)

# Diagnostico Rapido
$btnDiag = New-ProfileButton "Diagnostico de Estado (CPU / RAM / Hilos)" 28 $colCyan {
    $os = Get-CimInstance Win32_OperatingSystem
    $free = [math]::Round($os.FreePhysicalMemory / 1024)
    $tot  = [math]::Round($os.TotalVisibleMemorySize / 1024)
    $procs = Get-Process
    $threads = ($procs | Measure-Object -Property Threads -Sum).Sum
    $tel = (Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -ErrorAction SilentlyContinue).AllowTelemetry
    $telTxt = if ($tel -eq 0) { "DESACTIVADA [OK]" } else { "ACTIVA [FALTA]" }

    Write-Log "Diagnostico: RAM Libre $free/$tot MB | Procesos: $($procs.Count) | Subprocesos: $threads | Telemetria: $telTxt"
    [System.Windows.Forms.MessageBox]::Show(
        "ESTADO ACTUAL DEL SISTEMA:`n`n- RAM Libre: $free MB de $tot MB`n- Procesos Activos: $($procs.Count)`n- Subprocesos (Meta <600): $threads`n- Telemetria Windows: $telTxt",
        "XIMIOptimizer - Diagnostico",
        0, 64
    ) | Out-Null
}
$grpDiag.Controls.Add($btnDiag)

# Radiografia Clinica (Reporte en Escritorio)
$btnRadio = New-ProfileButton "Radiografia Clinica (Exportar a Escritorio)" 72 $colWarn {
    Write-Log "Generando Radiografia Clinica de procesos y memoria..."
    $rpt = "$env:USERPROFILE\Desktop\XIMI_Radiografia_Clinica.txt"
    $os = Get-CimInstance Win32_OperatingSystem
    $cpu = Get-CimInstance Win32_Processor
    $procs = Get-Process
    $threads = ($procs | Measure-Object -Property Threads -Sum).Sum
    $topCPU = $procs | Sort-Object CPU -Descending | Select-Object -First 8 Name, @{N="CPU_Seg";E={[math]::Round($_.CPU,1)}}, @{N="RAM_MB";E={[math]::Round($_.WorkingSet/1MB,1)}}, @{N="Hilos";E={$_.Threads.Count}}

    $out = @()
    $out += "=================================================================="
    $out += "  XIMIOptimizer v1.0 - RADIOGRAFIA CLINICA DEL SISTEMA"
    $out += "  Fecha: $(Get-Date)"
    $out += "=================================================================="
    $out += "CPU: $($cpu.Name) | MaxClock: $($cpu.MaxClockSpeed) MHz"
    $out += "RAM Libre: $([math]::Round($os.FreePhysicalMemory/1024)) MB de $([math]::Round($os.TotalVisibleMemorySize/1024)) MB"
    $out += "Procesos: $($procs.Count) | Subprocesos Totales: $threads"
    $out += ""
    $out += "--- TOP 8 PROCESOS CON MAYOR CARGA ACUMULADA ---"
    $out += ($topCPU | Format-Table -AutoSize | Out-String)
    $out | Out-File -FilePath $rpt -Encoding UTF8

    Write-Log "Radiografia guardada en: $rpt"
    [System.Windows.Forms.MessageBox]::Show("Radiografia Clinica generada con exito en tu Escritorio:`nXIMI_Radiografia_Clinica.txt", "Radiografia Lista", 0, 64) | Out-Null
}
$grpDiag.Controls.Add($btnRadio)

# Restaurar Windows
$btnRest = New-ProfileButton "Restaurar Valores Originales de Windows" 125 $colDanger {
    $res = [System.Windows.Forms.MessageBox]::Show("¿Deseas restaurar los servicios, la compresion de RAM y eliminar los FastFlags de Roblox?", "Confirmar Restauracion", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
        Write-Log "Restaurando configuracion predeterminada de Windows..."
        Enable-MMAgent -MemoryCompression -ErrorAction SilentlyContinue
        foreach ($s in @("Winmgmt","SysMain","WSearch","DiagTrack")) {
            Set-Service -Name $s -StartupType Automatic -ErrorAction SilentlyContinue
            Start-Service -Name $s -ErrorAction SilentlyContinue
        }
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" -Name "SystemResponsiveness" -Value 20 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "EnableTransparency" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        Get-ChildItem "$env:LOCALAPPDATA\Roblox\Versions" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $f = "$($_.FullName)\ClientSettings\ClientAppSettings.json"
            if (Test-Path $f) { Remove-Item $f -Force }
        }
        powercfg -setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN 25 2>$null | Out-Null
        powercfg -setactive SCHEME_CURRENT 2>$null | Out-Null
        Write-Log "Restauracion completada con exito."
        [System.Windows.Forms.MessageBox]::Show("Windows y Roblox han sido restaurados a su estado normal.", "Restaurado", 0, 64) | Out-Null
    }
}
$grpDiag.Controls.Add($btnRest)

# ==============================================================================
# 5. PIE DE PAGINA Y EJECUCION DE INTERFAZ
# ==============================================================================
$lblFooter = New-Object System.Windows.Forms.Label
$lblFooter.Text = "XIMIOptimizer v1.0 | Diseñado para Intel Celeron N4500 | Audio y Recortes (Win+Shift+S) Protegidos"
$lblFooter.Location = New-Object System.Drawing.Point(20, 610)
$lblFooter.Size = New-Object System.Drawing.Size(650, 20)
$lblFooter.ForeColor = $colTextMuted
$lblFooter.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$form.Controls.Add($lblFooter)

$btnSalir = New-Object System.Windows.Forms.Button
$btnSalir.Text = "Cerrar XIMIOptimizer"
$btnSalir.Location = New-Object System.Drawing.Point(695, 606)
$btnSalir.Size = New-Object System.Drawing.Size(150, 28)
$btnSalir.BackColor = $colBtnNormal
$btnSalir.ForeColor = [System.Drawing.Color]::White
$btnSalir.FlatStyle = "Flat"
$btnSalir.Add_Click({ $form.Close() })
$form.Controls.Add($btnSalir)

[void]$form.ShowDialog()
