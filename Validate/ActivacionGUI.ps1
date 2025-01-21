if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing

$urlIcono = "https://raw.githubusercontent.com/mggons93/Mggons/refs/heads/main/Validate/R.ico"
$rutaTemporalIcono = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "mi_icono.ico")

Invoke-WebRequest -Uri $urlIcono -OutFile $rutaTemporalIcono


function Verificar-Activacion {
    $salida = (slmgr -dli 2>&1 | Out-String)
    return $salida -match "Licensed"
}

if (Verificar-Activacion) {
    [System.Windows.MessageBox]::Show("Windows está activado")
    exit
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Activación de Windows by Mggons" Height="230" Width="400" Background="#778899">
    <Grid VerticalAlignment="Center" HorizontalAlignment="Center">
        <StackPanel Margin="10" HorizontalAlignment="Center" VerticalAlignment="Center">
            <TextBlock Text="Activación de Windows Ver: 1.1" FontSize="20" FontWeight="Bold" Margin="0,0,0,20" HorizontalAlignment="Center"/>
            <TextBlock Text="Seleccione el método de activación:" FontWeight="Bold" Margin="0,0,0,10" HorizontalAlignment="Center"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <RadioButton x:Name="SerialRadioButton" Content="Activación por Serial" Margin="0,0,20,0"/>
                <RadioButton x:Name="AutomaticaRadioButton" Content="Activación Automática"/>
            </StackPanel>
            <TextBox x:Name="SerialTextBox" Width="200" Height="25" Visibility="Collapsed" Text="Ingrese el serial aquí" Margin="0,10,0,0" HorizontalAlignment="Center"/>
            <Button Content="Activar" Width="100" Height="30" Margin="0,10,0,0" x:Name="ActivarButton" HorizontalAlignment="Center"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([System.Uri]::new($rutaTemporalIcono))

$serialRadioButton = $window.FindName("SerialRadioButton")
$automaticaRadioButton = $window.FindName("AutomaticaRadioButton")
$serialTextBox = $window.FindName("SerialTextBox")
$activarButton = $window.FindName("ActivarButton")

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
        $serial = $serialTextBox.Text
        $regex = "^[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$"
        if ($serial -match $regex) {
            $command = "slmgr -ipk $serial"
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -WindowStyle Hidden -Wait | Out-Null
            Start-Sleep -Seconds 5
            if (Verificar-Activacion) {
                [System.Windows.MessageBox]::Show("La licencia ha sido instalada y activada correctamente.")
            } else {
                #[System.Windows.MessageBox]::Show("Error al activar la licencia.")
            }
        } else {
            [System.Windows.MessageBox]::Show("Error: El serial no tiene el formato correcto o no coincide.")
        }
    } elseif ($automaticaRadioButton.IsChecked -eq $true) {
        $url = "https://raw.githubusercontent.com/%blank%massgravel/Microsoft-%blank%Activation-Scripts/refs/%blank%heads/master/MAS/All-In-%blank%One-Version-KL/MAS_AIO.%blank%cmd"
        $url = $url -replace "%blank%", ""
        $outputPath1 = "$env:TEMP\O%blank%hook_Acti%blank%vation_AI%blank%O.cmd"
        $outputPath1 = $outputPath1 -replace "%blank%", ""
        Invoke-WebRequest -Uri $url -OutFile $outputPath1
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $outputPath1 /HWID" -WindowStyle Hidden -Wait
        Remove-Item -Path $outputPath1 -Force
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

$window.ShowDialog()
Remove-Item -Path $rutaTemporalIcono -Force
