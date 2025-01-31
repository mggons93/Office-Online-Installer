function Test-Admin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    # Si no es administrador, reiniciar como administrador
    $scriptPath = $MyInvocation.MyCommand.Path
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    return  # Se usa 'return' para evitar que el script se detenga
}

# Enable TLSv1.2 for compatibility with older clients for current session
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Función para registrar mensajes en el log
function Add-LogMessage {
    param ([string]$Message)
    $logFile = "C:\OfficeSetupLog.txt"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    Add-Content -Path $logFile -Value $logMessage
}

# Verificar y establecer la ruta del directorio de Office
if (Test-Path "$env:ProgramFiles\Microsoft Office\Office16") {
    Set-Location "$env:ProgramFiles\Microsoft Office\Office16"
    Add-LogMessage "Cambiado al directorio: $env:ProgramFiles\Microsoft Office\Office16"
} elseif (Test-Path "$env:ProgramFiles(x86)\Microsoft Office\Office16") {
    Set-Location "$env:ProgramFiles(x86)\Microsoft Office\Office16"
    Add-LogMessage "Cambiado al directorio: $env:ProgramFiles(x86)\Microsoft Office\Office16"
} else {
    $errorMsg = "No se encontró la instalación de Microsoft Office en Office16."
    Add-LogMessage $errorMsg
    Write-Error $errorMsg
    exit
}

# Configuración de la tarea programada y script
$taskName = "RemoveOfficeLicense"

# 1. Solicitar al usuario que seleccione la edición de Office
Write-Host "Seleccione la edición de Office a activar:"
Write-Host "1. Office 2016"
Write-Host "2. Office 2019"
Write-Host "3. Office 2021"
Write-Host "4. Office 2024"

$editionChoice = Read-Host -Prompt "Ingrese el número correspondiente a la edición"

# Validar la entrada del usuario
switch ($editionChoice) {
    "1" { $edition = "2016"; $licensePattern = "ProPlusVL*.xrm-ms" }
    "2" { $edition = "2019"; $licensePattern = "ProPlus2019VL*.xrm-ms" }
    "3" { $edition = "2021"; $licensePattern = "ProPlus2021VL*.xrm-ms" }
    "4" { $edition = "2024"; $licensePattern = "ProPlus2024VL*.xrm-ms" }
    default {
        Write-Host "Selección no válida. El script terminará." -ForegroundColor Red
        Add-LogMessage "Selección no válida realizada."
        exit
    }
}

# 2. Solicitar al usuario que ingrese la clave de producto
$productKey = Read-Host -Prompt "Por favor, inserte la clave de producto de Office"

if (-not $productKey) {
    Write-Host "No se ingresó ninguna clave de producto. El script terminará." -ForegroundColor Red
    Add-LogMessage "No se ingresó ninguna clave de producto."
    exit
}

# 3. Configurar la clave por volumen
Write-Host "Configurando licencia de Office para la edición $edition, Por Favor Espere..." -ForegroundColor Cyan
Add-LogMessage "Configurando licencia de Office para la edición $edition con la clave ingresada."

# Establecer la ruta de licencias según la edición seleccionada
$licensesPath = "..\root\Licenses16"

# Obtener todos los archivos que coincidan con el patrón
$licenseFiles = Get-ChildItem -Path $licensesPath -Filter $licensePattern

# Verificar si se encontraron archivos
if ($licenseFiles.Count -eq 0) {
    Write-Host "No se encontraron archivos que coincidan con el patrón $licensePattern en $licensesPath" -ForegroundColor Yellow
} else {
    foreach ($licenseFile in $licenseFiles) {
        # Construir la ruta completa al archivo de licencia
        $licenseFilePath = Join-Path -Path $licensesPath -ChildPath $licenseFile.Name

        # Ejecutar ospp.vbs para instalar la licencia
        Start-Process -FilePath "cscript.exe" -ArgumentList "`"ospp.vbs`" /inslic:`"$licenseFilePath`"" -Wait -NoNewWindow | Out-Null
    }
    Write-Host "Instalación de licencias completada." -ForegroundColor Green
}

# 4. Activar la clave de producto ingresada
Start-Process -FilePath "cscript.exe" -ArgumentList "ospp.vbs /inpkey:$productKey" -Wait -NoNewWindow | Out-Null
Start-Process -FilePath "cscript.exe" -ArgumentList "ospp.vbs /act" -Wait -NoNewWindow | Out-Null

Write-Host "Licencia instalada y activada correctamente." -ForegroundColor Green
Add-LogMessage "Licencia instalada y activada correctamente."


# 5. Crear un script para eliminar la licencia
Write-Host "Creando script para eliminar la licencia..." -ForegroundColor Cyan
Add-LogMessage "Creando script para eliminar la licencia."

if (!(Test-Path -Path "$env:ProgramData\Scripts")) {
    New-Item -ItemType Directory -Path "$env:ProgramData\Scripts" | Out-Null
    Add-LogMessage "Directorio $env:ProgramData\Scripts creado."
}


$ScriptBat = @'
@echo off
:: Verificar si el script se está ejecutando con privilegios de administrador
NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Elevando permisos...
    :: Ejecutar este mismo archivo como administrador
    powershell -Command "Start-Process cmd -ArgumentList '/c %~s0' -Verb runAs"
    exit
)

:: Ejecutar el script PowerShell con privilegios elevados
powershell -NoProfile -ExecutionPolicy Bypass -File "%ProgramData%\Scripts\RemoveOfficeLicense.ps1"
exit
'@

# Utilizar Out-File para escribir el contenido al archivo
$ScriptBatPath = "$env:ProgramData\Scripts\RemoveOfficeLicense.bat"
$ScriptBat | Out-File -FilePath $ScriptBatPath -Encoding UTF8


# Crear el script de eliminación sin afectar la estructura interna
$removalScript = @'
# Verificar si PowerShell se está ejecutando con privilegios de administrador
$IsAdmin = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
$IsAdminRole = $IsAdmin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdminRole) {
    Write-Host "Este script debe ejecutarse como Administrador. Cerrando..." -ForegroundColor Red
    exit
}

$taskName = "RemoveOfficeLicense"

# Verificar si la tarea ya existe
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Write-Host "La tarea programada '$taskName' ya existe. Eliminándola..." -ForegroundColor Yellow
    Add-LogMessage "La tarea programada '$taskName' ya existía. Procediendo a eliminarla."
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
}
"@
$consolePtr = [Win32]::GetConsoleWindow()
[Win32]::ShowWindow($consolePtr, 0)  # Oculta la ventana (0: ocultar, 5: mostrar)

# Eliminar licencia de Office
# Verificar y establecer la ruta del directorio de Office
if (Test-Path "$env:ProgramFiles\Microsoft Office\Office16") {
    $officePath = "$env:ProgramFiles\Microsoft Office\Office16"
    Write-Host "Cambiado al directorio: $officePath" -ForegroundColor Yellow
} elseif (Test-Path "$env:ProgramFiles(x86)\Microsoft Office\Office16") {
    $officePath = "$env:ProgramFiles(x86)\Microsoft Office\Office16"
    Write-Host "Cambiado al directorio: $officePath" -ForegroundColor Yellow
} else {
    Write-Host "No se encontró la instalación de Microsoft Office en Office16." -ForegroundColor Red
    exit
}

Write-Host 'Eliminando licencia de Office...' -ForegroundColor Cyan

# Ejecutar el comando dstatus usando cscript
$osppStatus = cscript //NoLogo "$officePath\ospp.vbs" /dstatus

# Asegúrate de que la salida de ospp.vbs esté en un formato adecuado
$osppStatus = $osppStatus -join "`n"  # Para asegurar que la salida es una cadena de texto completa

# Extrae las claves de producto de la salida usando una expresión regular que busque los últimos 5 caracteres
$pattern = "Last 5 characters of installed product key: (\S{5})"
$matches = [regex]::Matches($osppStatus, $pattern)

# Recorre todas las claves de producto encontradas
foreach ($match in $matches) {
    # Obtiene los últimos 5 caracteres de la clave
    $last5Digits = $match.Groups[1].Value

    # Muestra el mensaje de eliminación de la clave de 5 dígitos
    Write-Host "Eliminando la clave de producto: $last5Digits" -ForegroundColor Green

    # Elimina la clave de producto
    cscript //NoLogo "$officePath\ospp.vbs" /unpkey:$last5Digits
}

Write-Host "Proceso completado. Las licencias han sido eliminadas." -ForegroundColor Cyan
'@

# Utilizar Out-File para escribir el contenido al archivo
$removalScriptPath = "$env:ProgramData\Scripts\RemoveOfficeLicense.ps1"
$removalScript | Out-File -FilePath $removalScriptPath -Encoding UTF8

Write-Host "Script creado en $removalScriptPath" -ForegroundColor Green
Add-LogMessage "Script de eliminación de licencia creado en $removalScriptPath."

# Verificar si la tarea ya existe antes de registrarla
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Write-Host "La tarea programada '$taskName' ya existe. Eliminándola..." -ForegroundColor Yellow
    Add-LogMessage "La tarea programada '$taskName' ya existía. Procediendo a eliminarla."
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# 6. Crear una tarea programada con privilegios elevados
Write-Host "Creando tarea programada para eliminar la licencia con privilegios elevados en un año..." -ForegroundColor Cyan
Add-LogMessage "Creando tarea programada con privilegios elevados para eliminar la licencia en un año."

$action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$env:ProgramData\Scripts\RemoveOfficeLicense.bat`""
$trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddYears(1))
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# 7. Registrar la tarea con privilegios elevados
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal | Out-Null

Write-Host "Tarea programada '$taskName' creada correctamente para ejecutarse con privilegios elevados en un año." -ForegroundColor Green
Add-LogMessage "Tarea programada '$taskName' creada correctamente con privilegios elevados."
