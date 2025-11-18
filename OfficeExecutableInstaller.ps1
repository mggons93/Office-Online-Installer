# Funcion para reiniciar el script con privilegios de administrador
function Start-ProcessAsAdmin {
    param (
        [string]$file,
        [string[]]$arguments = @()
    )
    Start-Process -FilePath $file -ArgumentList $arguments -Verb RunAs
}

# Comprobar si el script se está ejecutando como administrador
$scriptPath = $MyInvocation.MyCommand.Path
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-ProcessAsAdmin -file "powershell.exe" -arguments "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    exit
}

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# CONFIGURACIÓN
$owner = "mggons93"
$repo = "Office-Online-Installer"
$downloadFolder = "$env:TEMP\OfficeInstaller"

# Crear carpeta si no existe
if (-not (Test-Path $downloadFolder)) {
    New-Item -ItemType Directory -Path $downloadFolder | Out-Null
}

# Obtener información del último release
$releaseUrl = "https://api.github.com/repos/$owner/$repo/releases/latest"
$headers = @{ "User-Agent" = "$owner" }

try {
    $release = Invoke-RestMethod -Uri $releaseUrl -Headers $headers
} catch {
    Write-Error "No se pudo obtener el último release: $_"
    exit 1
}

# Buscar archivo .exe
$exeAsset = $release.assets | Where-Object { $_.name -like "*.exe" } | Select-Object -First 1

if (-not $exeAsset) {
    Write-Error "No se encontró ningún archivo .exe en el release más reciente."
    exit 1
}

$exeName = $exeAsset.name
$exeUrl = $exeAsset.browser_download_url
$localExePath = Join-Path $downloadFolder $exeName

# Descargar si no está ya
if (-not (Test-Path $localExePath)) {
    Write-Host "Descargando $exeName..."
    Invoke-WebRequest -Uri $exeUrl -OutFile $localExePath -Headers $headers
} else {
    Write-Host "El archivo ya está descargado."
}

# 🛡️ NUEVA EXCLUSIÓN — automática según el usuario
$newExtraExclusion = Join-Path $env:TEMP "Ohook_Activation_AIO.cmd"
$newCmdName = "Ohook_Activation_AIO.cmd"

try {
    $defender = Get-Service -Name WinDefend -ErrorAction SilentlyContinue
    if ($defender -and $defender.Status -eq "Running") {

        $mp = Get-MpPreference
        $excludedPaths = $mp.ExclusionPath
        $excludedProcesses = $mp.ExclusionProcess

        # Excluir instalador
        if ($excludedPaths -notcontains $localExePath) {
            Write-Host "Agregando exclusión de ruta..."
            Add-MpPreference -ExclusionPath $localExePath
        } else {
            Write-Host "Ruta ya excluida."
        }

        if ($excludedProcesses -notcontains $exeName) {
            Write-Host "Agregando exclusión de proceso..."
            Add-MpPreference -ExclusionProcess $exeName
        } else {
            Write-Host "Proceso ya excluido."
        }

        # 🔥 EXCLUSIÓN NUEVA del archivo TEMP
        if ($excludedPaths -notcontains $newExtraExclusion) {
            Write-Host "Agregando exclusión del archivo TEMP: $newExtraExclusion"
            Add-MpPreference -ExclusionPath $newExtraExclusion
        } else {
            Write-Host "La exclusión TEMP ya existe."
        }

        if ($excludedProcesses -notcontains $newCmdName) {
            Add-MpPreference -ExclusionProcess $newCmdName
            Write-Host "Proceso excluido: $newCmdName"
        }

    } else {
        Write-Warning "Windows Defender no está activo o no disponible."
    }
} catch {
    Write-Warning "No se pudo agregar exclusión a Windows Defender: $_"
}

# Ejecutar el instalador
Write-Host "Ejecutando $exeName..."
Start-Process -FilePath $localExePath -Wait

# Limpiar
try {
    Remove-Item -Path $localExePath -Force
    Write-Host "Instalador eliminado: $localExePath"

    if (Test-Path $downloadFolder) {
        Remove-Item -Path $downloadFolder -Recurse -Force
        Write-Host "Carpeta temporal eliminada: $downloadFolder"
    }
} catch {
    Write-Warning "No se pudo limpiar todo: $_"
}
