if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    
    Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Win32 {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll", ExactSpelling = true)]
    public static extern IntPtr GetConsoleWindow();
}
"@

$consolePtr = [Win32]::GetConsoleWindow()
[Win32]::ShowWindow($consolePtr, 6)

# Cargar las bibliotecas necesarias
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing

# Definir la URL del icono
$urlIcono = "https://raw.githubusercontent.com/mggons93/Mggons/refs/heads/main/Validate/R.ico"
$rutaTemporalIcono = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "mi_icono.ico")

# Descargar el icono desde la URL y guardarlo en la ubicación temporal
Invoke-WebRequest -Uri $urlIcono -OutFile $rutaTemporalIcono

# Función para verificar el estado de activación de Windows
function Verificar-Activacion {
    $salida = (slmgr -dli 2>&1 | Out-String)
    return $salida -match "Licensed"
}

# Verificar si Windows está activado
if (Verificar-Activacion) {
    [System.Windows.MessageBox]::Show("Windows está activado")
    exit
}

# Crear la ventana principal
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Activación de Windows by Mggons" Height="250" Width="400" Background="#778899">
    <Grid VerticalAlignment="Center" HorizontalAlignment="Center">
        <StackPanel Margin="10" HorizontalAlignment="Center" VerticalAlignment="Center">
            <TextBlock Text="Activación de Windows Ver: 1.1" FontSize="20" FontWeight="Bold" Margin="0,0,0,20" HorizontalAlignment="Center"/>
            <TextBlock Text="Seleccione el método de activación:" FontWeight="Bold" Margin="0,0,0,10" HorizontalAlignment="Center"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <RadioButton x:Name="SerialRadioButton" Content="Activación por Serial" Margin="0,0,20,0"/>
                <RadioButton x:Name="AutomaticaRadioButton" Content="Activación Automática"/>
            </StackPanel>
            <TextBox x:Name="SerialTextBox" Width="200" Height="25" Visibility="Collapsed" Text="Ingrese el serial aquí" Margin="0,10,0,0" HorizontalAlignment="Center"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <Button Content="Activar" Width="100" Height="30" Margin="0,10,0,0" x:Name="ActivarButton" HorizontalAlignment="Center"/>
                <Button Content="Cambiar Edición" Width="150" Height="30" Margin="20,10,0,0" x:Name="CambiarEdicionButton" HorizontalAlignment="Center"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>
"@

# Convertir el XAML en objetos de WPF
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Establecer el icono de la ventana
$window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([System.Uri]::new($rutaTemporalIcono))

# Obtener los elementos de la interfaz
$serialRadioButton = $window.FindName("SerialRadioButton")
$automaticaRadioButton = $window.FindName("AutomaticaRadioButton")
$serialTextBox = $window.FindName("SerialTextBox")
$activarButton = $window.FindName("ActivarButton")
$cambiarEdicionButton = $window.FindName("CambiarEdicionButton")

# Añadir los eventos
$serialRadioButton.Add_Checked({
    $serialTextBox.Visibility = 'Visible'
})

$automaticaRadioButton.Add_Checked({
    $serialTextBox.Visibility = 'Collapsed'
})

$serialTextBox.Add_GotFocus({
    if ($serialTextBox.Text -eq "Ingrese el serial aquí") {
        $serialTextBox.Text = ""
    }
})

$serialTextBox.Add_LostFocus({
    if ($serialTextBox.Text -eq "") {
        $serialTextBox.Text = "Ingrese el serial aquí"
    }
})

$activarButton.Add_Click({
    if ($serialRadioButton.IsChecked -eq $true) {
        # Código para activación por serial
        $serial = $serialTextBox.Text
        $regex = "^[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$"
        if ($serial -match $regex) {
            $command = "slmgr -ipk $serial"
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -WindowStyle Hidden -Wait | Out-Null
            # Verificar la activación
            Start-Sleep -Seconds 5
            if (Verificar-Activacion) {
                [System.Windows.MessageBox]::Show("La licencia ha sido instalada y activada correctamente.")
            } else {
                [System.Windows.MessageBox]::Show("Error al activar la licencia.")
            }
        } else {
            [System.Windows.MessageBox]::Show("Error: El serial no tiene el formato correcto o no coincide.")
        }
    } elseif ($automaticaRadioButton.IsChecked -eq $true) {
        # Código para activación automática
        $url = "https://raw.githubusercontent.com/%blank%massgravel/Microsoft-%blank%Activation-Scripts/refs/%blank%heads/master/MAS/All-In-%blank%One-Version-KL/MAS_AIO.%blank%cmd"
        $url = $url -replace "%blank%", ""
        $outputPath1 = "$env:TEMP\O%blank%hook_Acti%blank%vation_AI%blank%O.cmd"
        $outputPath1 = $outputPath1 -replace "%blank%", ""
        Invoke-WebRequest -Uri $url -OutFile $outputPath1
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $outputPath1 /HWID" -WindowStyle Hidden -Wait
        Remove-Item -Path $outputPath1 -Force
        # Verificar la activación después de la activación automática
        Start-Sleep -Seconds 5
        if (Verificar-Activacion) {
            [System.Windows.MessageBox]::Show("Windows ha sido activado correctamente.")
        } else {
            #[System.Windows.MessageBox]::Show("Error al activar Windows.")
        }
    } else {
        [System.Windows.MessageBox]::Show("Seleccione un método de activación.")
    }
})

$cambiarEdicionButton.Add_Click({
    # Código para cambiar la edición de Windows
$url = "https://raw.github%blank%usercontent.com/massgravel/Microsoft-Acti%blank%vation-Scripts/refs/h%blank%eads/master/MAS/Sepa%blank%rate-Files-Version/Change_W%blank%indows_Edi%blank%tion.cmd"
$url = $url -replace "%blank%", ""  # Reemplazar los '%blank%' con una cadena vacía
$outputPath2 = "$env:TEMP\Change_Windows_Edi%blank%tion.cmd"
Invoke-WebRequest -Uri $url -OutFile $outputPath2
Start-Process -FilePath "cmd.exe" -ArgumentList "/c $outputPath2" -Wait
Remove-Item -Path $outputPath2 -Force
    [System.Windows.MessageBox]::Show("La edición de Windows ha sido cambiada.")
})

# Mostrar la ventana y eliminar el icono al cerrar
$window.ShowDialog()
Remove-Item -Path $rutaTemporalIcono -Force

[Win32]::ShowWindow($consolePtr, 9) # 9 = Restaurar la ventana
