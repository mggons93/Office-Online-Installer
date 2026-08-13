function Start-ProcessAsAdmin {
    param (
        [string]$file,
        [string[]]$arguments = @()
    )
    Start-Process -FilePath $file -ArgumentList $arguments -Verb RunAs
}

$scriptPath = $MyInvocation.MyCommand.Path
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {

    Start-ProcessAsAdmin -file "powershell.exe" `
        -arguments "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    exit
}

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ================= CONFIGURACIÓN =================
$owner = "mggons93"
$repo  = "Office-Online-Installer"
$downloadFolder = "$env:TEMP\OfficeInstaller"

# Detectar arquitectura del SO
$is64 = [Environment]::Is64BitOperatingSystem
$arch = if ($is64) { "x64" } else { "x86" }

Write-Host "Sistema detectado: $arch"

# Crear carpeta temporal
if (-not (Test-Path $downloadFolder)) {
    New-Item -ItemType Directory -Path $downloadFolder | Out-Null
}

# ================= OBTENER RELEASE =================
$releaseUrl = "https://api.github.com/repos/$owner/$repo/releases/latest"
$headers = @{ "User-Agent" = "$owner" }

try {
    $release = Invoke-RestMethod -Uri $releaseUrl -Headers $headers
} catch {
    Write-Error "No se pudo obtener el último release: $_"
    exit 1
}

# ================= SELECCIÓN x86 / x64 =================
$exeAsset = $release.assets | Where-Object {
    $_.name -match "\.exe$" -and $_.name -match $arch
} | Select-Object -First 1

# Fallback si no encuentra coincidencia exacta
if (-not $exeAsset) {
    Write-Warning "No se encontró instalador específico $arch, usando el primero disponible."
    $exeAsset = $release.assets | Where-Object {
        $_.name -match "\.exe$"
    } | Select-Object -First 1
}

if (-not $exeAsset) {
    Write-Error "No se encontró ningún archivo .exe en el release."
    exit 1
}

$exeName = $exeAsset.name
$exeUrl  = $exeAsset.browser_download_url
$localExePath = Join-Path $downloadFolder $exeName

# ================= DESCARGA =================
if (-not (Test-Path $localExePath)) {
    Write-Host "Descargando $exeName..."
    Invoke-WebRequest -Uri $exeUrl -OutFile $localExePath -Headers $headers
} else {
    Write-Host "El archivo ya está descargado."
}

# ================= EXCLUSIONES DEFENDER =================
$newExtraExclusion = Join-Path $env:TEMP "Ohook_Activation_AIO.cmd"
$newCmdName = "Ohook_Activation_AIO.cmd"

try {
    $defender = Get-Service -Name WinDefend -ErrorAction SilentlyContinue
    if ($defender -and $defender.Status -eq "Running") {

        $mp = Get-MpPreference

        if ($mp.ExclusionPath -notcontains $localExePath) {
            Add-MpPreference -ExclusionPath $localExePath
        }

        if ($mp.ExclusionProcess -notcontains $exeName) {
            Add-MpPreference -ExclusionProcess $exeName
        }

        if ($mp.ExclusionPath -notcontains $newExtraExclusion) {
            Add-MpPreference -ExclusionPath $newExtraExclusion
        }

        if ($mp.ExclusionProcess -notcontains $newCmdName) {
            Add-MpPreference -ExclusionProcess $newCmdName
        }
    }
} catch {
    Write-Warning "No se pudieron aplicar exclusiones: $_"
}

# ================= EJECUCIÓN =================
Write-Host "Ejecutando $exeName..." -Wait

# ================= LIMPIEZA =================
try {
    Remove-Item -Path $downloadFolder -Recurse -Force
    Write-Host "Limpieza completada."
} catch {
    Write-Warning "No se pudo limpiar completamente: $_"
}
