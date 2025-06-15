# Funcion para reiniciar el script con privilegios de administrador
function Start-ProcessAsAdmin {
    param (
        [string]$file,
        [string[]]$arguments = @()
    )
    Start-Process -FilePath $file -ArgumentList $arguments -Verb RunAs
}

# Comprobar si el script se estiÂ¡ ejecutando como administrador
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    # Si no estiÂ¡ ejecutiÂ¡ndose como administrador, relanza el script con privilegios elevados
    Start-ProcessAsAdmin -file "powershell.exe" -arguments "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    exit
}


$ErrorActionPreference = "Stop"
# Enable TLSv1.2 for compatibility with older clients for current session
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# URL de la imagen de Office 365
$imageUrl = "https://granikos.eu/wp-content/uploads/2023/02/Logo-Microsoft365.png"

# Obtener la hora actual formateada para el nombre del archivo
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

# Ruta temporal con nombre diniÂ¡mico
$tempImagePath = [System.IO.Path]::Combine($env:TEMP, "office365_$timestamp.png")

# Descargar la imagen si no existe o si quieres forzar que se actualice cada vez
Invoke-WebRequest -Uri $imageUrl -OutFile $tempImagePath -UseBasicParsing

Add-Type -AssemblyName PresentationFramework
$is64Bit = [Environment]::Is64BitOperatingSystem

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

# Definir la URL del icono
$urlIcono = "https://raw.githubusercontent.com/mggons93/Mggons/refs/heads/main/OfficeIco.ico"
$rutaTemporalIcono = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "mi_icono.ico")

# Descargar el icono desde la URL y guardarlo en la ubicacion temporal
Invoke-WebRequest -Uri $urlIcono -OutFile $rutaTemporalIcono

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft Office Online By Mggons" Height="379" Width="470" Background="#8e77ab" ResizeMode="NoResize">
    <Grid Height="347">
        <!-- Fila 1 -->
        <TextBlock Text="Selecciona tu edicion de Office:" 
        HorizontalAlignment="Left" 
        VerticalAlignment="Top" 
        Margin="29,19,0,0" 
        FontWeight="Bold"/>
        <ComboBox x:Name="variantComboBox" HorizontalAlignment="Left" VerticalAlignment="Top" Width="200" Margin="10,42,0,0">
            <!-- Office 365 -->
            <ComboBoxItem IsEnabled="False">
                <TextBlock Text="--- Microsoft 365 ---" FontWeight="Bold"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="O365ProPlusRetail">
                <TextBlock Text="Office 365 Pro Plus"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="O365ProPlusEEANoTeamsRetail">
                <TextBlock Text="Office 365 Pro Plus (No Teams)"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="O365BusinessRetail">
                <TextBlock Text="Office 365 Business"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="O365BusinessEEANoTeamsRetail">
                <TextBlock Text="Office 365 Business (No Teams)"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="O365EduCloudRetail">
                <TextBlock Text="Office 365 Education Cloud"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="O365HomePremRetail">
                <TextBlock Text="Office 365 Home Premium"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="O365SmallBusPremRetail">
                <TextBlock Text="Office 365 Small Business Premium"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="ProjectStdRetail">
                <TextBlock Text="Office 365 Project Standard"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="ProjectProRetail">
                <TextBlock Text="Office 365 Project Professional"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="VisioStdRetail">
                <TextBlock Text="Office 365 Visio Standard"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="VisioProRetail">
                <TextBlock Text="Office 365 Visio Professional"/>
            </ComboBoxItem>

            <!-- Separador -->
            <ComboBoxItem IsEnabled="False">
                <TextBlock Text="--- Microsoft Office 2024 ---" FontWeight="Bold"/>
            </ComboBoxItem>

            <!-- Edicion 2024 -->
            <ComboBoxItem Tag="ProPlus2024Retail">
                <TextBlock Text="Office Professional Plus 2024"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="HomeBusiness2024Retail">
                <TextBlock Text="Home &amp; Business 2024"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="Home2024Retail">
                <TextBlock Text="Office Home 2024"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="ProjectStd2024Retail">
                <TextBlock Text="Project Standard 2024"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="ProjectPro2024Retail">
                <TextBlock Text="Project Professional 2024"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="VisioStd2024Retail">
                <TextBlock Text="Visio Standard 2024"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="VisioPro2024Retail">
                <TextBlock Text="Visio Professional 2024"/>
            </ComboBoxItem>

            <!-- Separador -->
            <ComboBoxItem IsEnabled="False">
                <TextBlock Text="--- Microsoft Office 2021 ---" FontWeight="Bold"/>
            </ComboBoxItem>

            <!-- Edicion 2021 -->
            <ComboBoxItem Tag="ProPlus2021Retail">
                <TextBlock Text="Office Professional Plus 2021"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="HomeBusiness2021Retail">
                <TextBlock Text="Office Home &amp; Business 2021"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="HomeStudent2021Retail">
                <TextBlock Text="Office Home &amp; Student 2021"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="Standard2021Retail">
                <TextBlock Text="Office Standard 2021"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="ProjectStd2021Retail">
                <TextBlock Text="Project Standard 2021"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="ProjectPro2021Retail">
                <TextBlock Text="Project Professional 2021"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="VisioStd2021Retail">
                <TextBlock Text="Visio Standard 2021"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="VisioPro2021Retail">
                <TextBlock Text="Visio Professional 2021"/>
            </ComboBoxItem>

            <!-- Separador -->
            <ComboBoxItem IsEnabled="False">
                <TextBlock Text="--- Microsoft Office  2019 ---" FontWeight="Bold"/>
            </ComboBoxItem>

            <!-- Edicion 2019 -->
            <ComboBoxItem Tag="ProPlus2019Retail">
                <TextBlock Text="Office Professional Plus 2019"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="HomeBusiness2019Retail">
                <TextBlock Text="Office Home &amp; Business 2019"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="HomeStudent2019Retail">
                <TextBlock Text="Office Home &amp; Student 2019"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="Standard2019Retail">
                <TextBlock Text="Office Standard 2019"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="ProjectStd2019Retail">
                <TextBlock Text="Project Standard 2019"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="ProjectPro2019Retail">
                <TextBlock Text="Project Professional 2019"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="VisioStd2019Retail">
                <TextBlock Text="Visio Standard 2019"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="VisioPro2019Retail">
                <TextBlock Text="Visio Professional 2019"/>
            </ComboBoxItem>

            <!-- Separador -->
            <ComboBoxItem IsEnabled="False">
                <TextBlock Text="--- Microsoft Office 2016 ---" FontWeight="Bold"/>
            </ComboBoxItem>

            <!-- Edicion 2016 -->
            <ComboBoxItem Tag="ProPlusRetail">
                <TextBlock Text="Office Professional Plus 2016"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="HomeBusinessRetail">
                <TextBlock Text="Office Home &amp; Business 2016"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="HomeStudentRetail">
                <TextBlock Text="Office Home &amp; Student 2016"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="StandardRetail">
                <TextBlock Text="Office Standard 2016"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="ProjectStdRetail">
                <TextBlock Text="Project Standard 2016"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="ProjectProRetail">
                <TextBlock Text="Project Professional 2016"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="VisioStdRetail">
                <TextBlock Text="Visio Standard 2016"/>
            </ComboBoxItem>
            <ComboBoxItem Tag="VisioProRetail">
                <TextBlock Text="Visio Professional 2016"/>
            </ComboBoxItem>
        </ComboBox>
        <CheckBox x:Name="architectureCheckBox" Content="OS x64" FontWeight="Bold" IsChecked="$($is64Bit)" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="221,45,0,0" IsEnabled="True"/>

        <!-- Fila 2 -->
        <TextBlock Text="Seleccion de Idioma:" 
        HorizontalAlignment="Left" 
        VerticalAlignment="Top" 
        Margin="53,69,0,0" 
        FontWeight="Bold"/>
        <ComboBox x:Name="languageComboBox" HorizontalAlignment="Left" VerticalAlignment="Top" Width="200" Margin="10,90,0,0">
            <ComboBoxItem Content="English" Tag="en-US"/>
            <ComboBoxItem Content="Arabic" Tag="ar-SA"/>
            <ComboBoxItem Content="Bulgarian" Tag="bg-BG"/>
            <ComboBoxItem Content="Czech" Tag="cs-CZ"/>
            <ComboBoxItem Content="Danish" Tag="da-DK"/>
            <ComboBoxItem Content="German" Tag="de-DE"/>
            <ComboBoxItem Content="Greek" Tag="el-GR"/>
            <ComboBoxItem Content="English UK" Tag="en-GB"/>
            <ComboBoxItem Content="Spanish" Tag="es-ES"/>
            <ComboBoxItem Content="Spanish Mexico" Tag="es-MX"/>
            <ComboBoxItem Content="Estonian" Tag="et-EE"/>
            <ComboBoxItem Content="Finnish" Tag="fi-FI"/>
            <ComboBoxItem Content="French Canada" Tag="fr-CA"/>
            <ComboBoxItem Content="French" Tag="fr-FR"/>
            <ComboBoxItem Content="Hebrew" Tag="he-IL"/>
            <ComboBoxItem Content="Hindi" Tag="hi-IN"/>
            <ComboBoxItem Content="Croatian" Tag="hr-HR"/>
            <ComboBoxItem Content="Hungarian" Tag="hu-HU"/>
            <ComboBoxItem Content="Indonesian" Tag="id-ID"/>
            <ComboBoxItem Content="Italian" Tag="it-IT"/>
            <ComboBoxItem Content="Japanese" Tag="ja-JP"/>
            <ComboBoxItem Content="Kazakh" Tag="kk-KZ"/>
            <ComboBoxItem Content="Korean" Tag="ko-KR"/>
            <ComboBoxItem Content="Lithuanian" Tag="lt-LT"/>
            <ComboBoxItem Content="Latvian" Tag="lv-LV"/>
            <ComboBoxItem Content="Malay (Latin)" Tag="ms-MY"/>
            <ComboBoxItem Content="Norwegian Bokmal" Tag="nb-NO"/>
            <ComboBoxItem Content="Dutch" Tag="nl-NL"/>
            <ComboBoxItem Content="Polish" Tag="pl-PL"/>
            <ComboBoxItem Content="Portuguese (Brazil)" Tag="pt-BR"/>
            <ComboBoxItem Content="Portuguese (Portugal)" Tag="pt-PT"/>
            <ComboBoxItem Content="Romanian" Tag="ro-RO"/>
            <ComboBoxItem Content="Russian" Tag="ru-RU"/>
            <ComboBoxItem Content="Slovak" Tag="sk-SK"/>
            <ComboBoxItem Content="Slovenian" Tag="sl-SI"/>
            <ComboBoxItem Content="Swedish" Tag="sv-SE"/>
            <ComboBoxItem Content="Thai" Tag="th-TH"/>
            <ComboBoxItem Content="Turkish" Tag="tr-TR"/>
            <ComboBoxItem Content="Ukrainian" Tag="uk-UA"/>
            <ComboBoxItem Content="Chinese (Simplified)" Tag="zh-CN"/>
            <ComboBoxItem Content="Chinese (Traditional)" Tag="zh-TW"/>
            <!-- Agrega aquiÂ­ miÂ¡s opciones si es necesario -->
        </ComboBox>
        <CheckBox x:Name="vlActivationCheckBox" 
        Content="Volumen License Code" 
        HorizontalAlignment="Left" 
        VerticalAlignment="Top" 
        Margin="220,93,0,0" FontWeight="Bold"/>

        <!-- Opciones de Activacion VL solo si se activa VL -->
        <StackPanel x:Name="vlOptionsPanel" Visibility="Collapsed" 
        HorizontalAlignment="Left" VerticalAlignment="Top" 
        Margin="10,125,0,0" Orientation="Horizontal">
            <StackPanel Orientation="Vertical" Margin="0,0,10,0">
                <TextBlock Text="Seleccione la edicion por Volumen:" HorizontalAlignment="Center" Margin="0,0,0,5"  FontWeight="Bold"/>
                <ComboBox x:Name="editionComboBox" HorizontalAlignment="Left" Width="200">
                    <!-- Separador -->
                    <ComboBoxItem IsEnabled="False">
                        <TextBlock Text="--- Microsoft Office  2016 ---" FontWeight="Bold"/>
                    </ComboBoxItem>
                    <ComboBoxItem Content="Office 2016 Standard VL"/>
                    <ComboBoxItem Content="Office 2016 Professional VL"/>
                    <ComboBoxItem Content="Project 2016 Standard VL"/>
                    <ComboBoxItem Content="Project 2016 Professional VL"/>
                    <ComboBoxItem Content="Visio 2016 Standard VL"/>
                    <ComboBoxItem Content="Visio 2016 Professional VL"/>
                    <!-- Separador -->
                    <ComboBoxItem IsEnabled="False">
                        <TextBlock Text="--- Microsoft Office  2019 ---" FontWeight="Bold"/>
                    </ComboBoxItem>
                    <ComboBoxItem Content="Office 2019 Standard VL"/>
                    <ComboBoxItem Content="Office 2019 Professional VL"/>
                    <ComboBoxItem Content="Project 2019 Standard VL"/>
                    <ComboBoxItem Content="Project 2019 Professional VL"/>
                    <ComboBoxItem Content="Visio 2019 Standard VL"/>
                    <ComboBoxItem Content="Visio 2019 Professional VL"/>
                    <!-- Separador -->
                    <ComboBoxItem IsEnabled="False">
                        <TextBlock Text="--- Microsoft Office  2021 ---" FontWeight="Bold"/>
                    </ComboBoxItem>
                    <ComboBoxItem Content="Office 2021 Standard VL"/>
                    <ComboBoxItem Content="Office 2021 Professional VL"/>
                    <ComboBoxItem Content="Project 2021 Standard VL"/>
                    <ComboBoxItem Content="Project 2021 Professional VL"/>
                    <ComboBoxItem Content="Visio 2021 Standard VL"/>
                    <ComboBoxItem Content="Visio 2021 Professional VL"/>
                    <!-- Separador -->
                    <ComboBoxItem IsEnabled="False">
                        <TextBlock Text="--- Microsoft Office  2024 ---" FontWeight="Bold"/>
                    </ComboBoxItem>
                    <ComboBoxItem Content="Office 2024 Standard VL"/>
                    <ComboBoxItem Content="Office 2024 Professional VL"/>
                    <ComboBoxItem Content="Project 2024 Standard VL"/>
                    <ComboBoxItem Content="Project 2024 Professional VL"/>
                    <ComboBoxItem Content="Visio 2024 Standard VL"/>
                    <ComboBoxItem Content="Visio 2024 Professional VL"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Orientation="Vertical">
                <TextBlock Text="Ingrese la clave de licencia:" HorizontalAlignment="Center" Margin="0,0,0,8" FontWeight="Bold"/>
                <TextBox x:Name="licenseKeyTextBox" HorizontalAlignment="Left" Width="216"/>
            </StackPanel>
        </StackPanel>

        <!-- Boton de instalacion y log -->
        <Button x:Name="installButton" Content="Instalar" HorizontalAlignment="Left" VerticalAlignment="Top" Width="80" Height="30" Margin="10,184,0,0" FontWeight="Bold">
            <Button.ToolTip>
                <ToolTip Content="Esta opcion permite instalar office con: Activacion por Volumen, Activacion Automatica o Sin Activacion." />
            </Button.ToolTip>
        </Button>

        <CheckBox x:Name="autoActivationCheckBox" Content="Activar Automatico" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="92,192,0,0" FontWeight="Bold" IsChecked="True"/>

       <Button x:Name="Activeread" Content="Solo Activar" HorizontalAlignment="Right" VerticalAlignment="Top" Width="96" Height="30" Margin="0,184,129,0" FontWeight="Bold">
            <Button.ToolTip>
                <ToolTip Content="Esta opcion es solo para activar,(solo si tienes office instalado previamente.)" />
            </Button.ToolTip>
        </Button>

        <Button x:Name="Donate" Content="Donate" HorizontalAlignment="Left" VerticalAlignment="Top" Width="93" Height="30" Margin="343,184,0,0" FontWeight="Bold">
            <Button.ToolTip>
                <ToolTip Content="Aqui nos puede donar y apotar en el proyecto." />
            </Button.ToolTip>
        </Button>

        <TextBlock Text="Informacion de la Instalacion:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="145,218,0,0" FontWeight="Bold"/>
        <TextBox x:Name="logTextBox" HorizontalAlignment="Left" VerticalAlignment="Top" Width="426" Height="60" Margin="10,241,0,0" IsReadOnly="True" VerticalScrollBarVisibility="Auto" FontWeight="Bold"/>
        <Label Content="Version de ODT: 3.2 " Height="27" HorizontalAlignment="Left" Margin="168,307,0,0" Name="label1" VerticalAlignment="Top" Width="124" FontWeight="Bold"/>
        <Image Height="73" HorizontalAlignment="Left" Margin="369,12,0,0" x:Name="image1" Stretch="Fill" VerticalAlignment="Top" Width="67" />
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Cargar la imagen en el control 'image1'
$imageControl = $window.FindName("image1")
$imageControl.Source = [System.Windows.Media.Imaging.BitmapImage]::new([Uri]::new($tempImagePath))
$variantComboBox = $window.FindName("variantComboBox")
$languageComboBox = $window.FindName("languageComboBox")
$vlActivationCheckBox = $window.FindName("vlActivationCheckBox")
$architectureCheckBox = $window.FindName("architectureCheckBox")
$editionComboBox = $window.FindName("editionComboBox")
$licenseKeyTextBox = $window.FindName("licenseKeyTextBox")
$installButton = $window.FindName("installButton")
$donateButton = $window.FindName("Donate")
$activereadButton = $window.FindName("Activeread")
$logTextBox = $window.FindName("logTextBox")
$vlOptionsPanel = $window.FindName("vlOptionsPanel")
$autoActivationCheckBox = $window.FindName("autoActivationCheckBox")


function Add-LogMessage {
    param (
        [string]$message
    )
    $logTextBox.Text += "$message`r`n"
}

$vlActivationCheckBox.Add_Checked({
    $vlOptionsPanel.Visibility = 'Visible'
    $autoActivationCheckBox.IsChecked = $false
    $autoActivationCheckBox.IsEnabled = $false
})

$vlActivationCheckBox.Add_Unchecked({
    $vlOptionsPanel.Visibility = 'Collapsed'
    $autoActivationCheckBox.IsEnabled = $true
    $autoActivationCheckBox.IsChecked = $false
})

$autoActivationCheckBox.Add_Checked({
    $vlActivationCheckBox.IsChecked = $false
    $autoActivationCheckBox.IsEnabled = $true
})

$autoActivationCheckBox.Add_Checked({
    $vlActivationCheckBox.IsChecked = $false
    $autoActivationCheckBox.IsEnabled = $false
})

$autoActivationCheckBox.Add_Unchecked({
    if (-not $vlActivationCheckBox.IsChecked) {
        $autoActivationCheckBox.IsChecked = $false
    }
})

$donateButton.Add_Click({
Start-Process "https://cutt.ly/DonacionSyA"
})

$activereadButton.Add_Click({
    $installButton.IsEnabled = $false
    # Desactivar el boton para evitar miltiples clics
    $activereadButton.IsEnabled = $false
    $activereadButton.Content = "Activando..."

    # Preparar las variables
    $url = "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/refs/heads/master/MAS/All-In-One-Version-KL/MAS_AIO.cmd"
    $outputPath1 = "$env:TEMP\Ohook_Activation_AIO.cmd"

    # Iniciar la tarea pesada en segundo plano
    $job = Start-Job -ScriptBlock {
        param($url, $outputPath1)
        $log = @()
        $log += "Iniciando Activacion. Espere..."
        try {
            Invoke-WebRequest -Uri $url -OutFile $outputPath1 -ErrorAction Stop

            $log += "Activando..."
            Start-Process -FilePath $outputPath1 -ArgumentList "/Ohook" -WindowStyle Hidden -Wait -Verb RunAs

            Remove-Item -Path $outputPath1 -Force
            $log += "Activacion completada."
        } catch {
            $log += "Error durante la activacion: $_"
        }
        return $log
    } -ArgumentList $url, $outputPath1

    # Monitorea el job y actualiza la interfaz
    while ($job.State -eq 'Running') {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
    }
    $msgs = Receive-Job $job
    foreach ($msg in $msgs) { Add-LogMessage $msg }
    Remove-Job $job

    # Restaurar el boton
    $installButton.IsEnabled = $true
    $activereadButton.IsEnabled = $true
    $activereadButton.Content = "Solo Activar"    
})

$installButton.Add_Click({

# Desactivar el boton para evitar multiples clics
    $installButton.IsEnabled = $false
    
    # Desactivar el boton para evitar multiples clics
    $activereadButton.IsEnabled = $false
    

    # Recolecta datos de la GUI
    $selectedVariant = $variantComboBox.SelectedItem.Tag
    $selectedLanguageFull = $languageComboBox.SelectedItem.Tag
    $selectedLanguage = $selectedLanguageFull -replace '.*\[(.*?)\]', '$1'
    $architecture = if ($architectureCheckBox.IsChecked) { "x64" } else { "x86" }
    $useVL = $vlActivationCheckBox.IsChecked
    $editionVL = if ($useVL) { $editionComboBox.SelectedItem.Content } else { $null }
    $licenseKey = if ($useVL) { $licenseKeyTextBox.Text } else { $null }
    $autoActivate = $autoActivationCheckBox.IsChecked

    $url = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=$selectedVariant&platform=$architecture&language=$selectedLanguage&version=O16GA"
    $outputFile = "$env:TEMP\${selectedVariant}-${selectedLanguage}-${architecture}.exe"
    $installButton.Content = "Instalando..." 
    # Mensaje de confirmacion
    $message = "Se procedera a descargar e instalar la siguiente variante de Office 365. ¿Desea continuar?"
    if ($useVL) {
        $message += "`n- Edicion VL: $editionVL"
        $message += "`n- Clave de licencia: $licenseKey"
    }
    $message += "`n- Variante: $selectedVariant"
    $message += "`n- Idioma: $selectedLanguage"
    $message += "`n- Sistema: $architecture"
    $message += "`n- Activacion Automatica: $($autoActivate -eq $true)"
    $result = [System.Windows.MessageBox]::Show($message, "Confirmar Instalacion", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)

    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        # Inicia la tarea pesada en segundo plano
        $job = Start-Job -ScriptBlock {
            param (
                $url, $outputFile, $autoActivate, $useVL, $editionVL, $licenseKey
            )
            $log = @()
            
            $log += "Descargando $url ..."
            try {
                Invoke-WebRequest -Uri $url -OutFile $outputFile -ErrorAction Stop
                $log += "Descarga completada. Iniciando la instalacion..."
                Start-Process -FilePath $outputFile -Wait
                "taskkill /f /im OfficeC2RClient.exe" | cmd
                $log += "Instalacion completada."
            } catch {
                $log += "Error durante la descarga o instalacion: $_"
                return $log
            }

            if ($autoActivate -and -not $useVL) {
                $log += "Iniciando Activacion. Espere..."
                $actUrl = "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/refs/heads/master/MAS/All-In-One-Version-KL/MAS_AIO.cmd"
                $outputPath1 = "$env:TEMP\Ohook_Activation_AIO.cmd"
                try {
                    Invoke-WebRequest -Uri $actUrl -OutFile $outputPath1 -ErrorAction Stop
                    $log += "Activando..."
                    Start-Process -FilePath $outputPath1 -ArgumentList "/Ohook" -WindowStyle Hidden -Wait -Verb RunAs
                    Remove-Item -Path $outputPath1 -Force
                    $log += "Activacion completada."
                } catch {
                    $log += "Error durante la activacion automitica: $_"
                }
            }

            if ($useVL) {
                $log += "Convirtiendo Office Retail a Volumen. Espere..."
                $edition = switch ($editionVL) {
                    "Office 2016 Standard VL" { "StandardVL" }
                    "Office 2016 Professional VL" { "ProPlusVL" }
                    "Project 2016 Standard VL" { "ProjectStdVL" }
                    "Project 2016 Professional VL" { "ProjectProVL" }
                    "Visio 2016 Standard VL" { "VisioStdVL" }
                    "Visio 2016 Professional VL" { "VisioProVL" }
                    "Office 2019 Standard VL" { "Standard2019VL" }
                    "Office 2019 Professional VL" { "ProPlus2019VL" }
                    "Project 2019 Standard VL" { "ProjectStd2019VL" }
                    "Project 2019 Professional VL" { "ProjectPro2019VL" }
                    "Visio 2019 Standard VL" { "VisioStd2019VL" }
                    "Visio 2019 Professional VL" { "VisioPro2019VL" }
                    "Office 2021 Standard VL" { "Standard2021VL" }
                    "Office 2021 Professional VL" { "ProPlus2021VL" }
                    "Project 2021 Standard VL" { "ProjectStd2021VL" }
                    "Project 2021 Professional VL" { "ProjectPro2021VL" }
                    "Visio 2021 Standard VL" { "VisioStd2021VL" }
                    "Visio 2021 Professional VL" { "VisioPro2021VL" }
                    "Office 2024 Standard VL" { "Standard2024VL" }
                    "Office 2024 Professional VL" { "ProPlus2024VL" }
                    "Project 2024 Standard VL" { "ProjectStd2024VL" }
                    "Project 2024 Professional VL" { "ProjectPro2024VL" }
                    "Visio 2024 Standard VL" { "VisioStd2024VL" }
                    "Visio 2024 Professional VL" { "VisioPro2024VL" }
                    default { $null }
                }
                $log += "Edicion seleccionada: $editionVL ($edition)"
                $officePath = if (Test-Path "$env:ProgramFiles\Microsoft Office\Office16") {
                    "$env:ProgramFiles\Microsoft Office\Office16"
                } elseif (Test-Path "$env:ProgramFiles(x86)\Microsoft Office\Office16") {
                    "$env:ProgramFiles(x86)\Microsoft Office\Office16"
                } else {
                    $log += "No se encontro la instalacion de Microsoft Office en Office16."
                    return $log
                }
                Set-Location $officePath
                $licenseFiles = Get-ChildItem -Path "..\root\Licenses16\" -Filter "$edition*.xrm-ms"
                $log += "Archivos de licencia encontrados: $($licenseFiles.Count)"
                foreach ($file in $licenseFiles) {
                    $filePath = "..\root\Licenses16\$($file.Name)"
                    $log += "Instalando Archivo: $filePath"
                    cscript ospp.vbs /inslic:"$filePath"
                }
                $log += "Instalando licencia: $licenseKey"
                cscript ospp.vbs /inpkey:$licenseKey
                cscript ospp.vbs /act
                $log += "Activacion completada."
            }
            return $log
        } -ArgumentList $url, $outputFile, $autoActivate, $useVL, $editionVL, $licenseKey

        # Monitorea el job y actualiza la interfaz
        while ($job.State -eq 'Running') {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 200
        }
        $msgs = Receive-Job $job
        foreach ($msg in $msgs) { Add-LogMessage $msg }
        Remove-Job $job
} else {
    # Codigo a ejecutar si el usuario selecciona "No"
    $installButton.IsEnabled = $true
    
    # Desactivar el boton para evitar multiples clics
    $activereadButton.IsEnabled = $true
    $installButton.Content = "Instalar"
    # Puedes agregar aqui cualquier otra accion que desees
}
})

$window.ShowDialog() | Out-Null
[Win32]::ShowWindow($consolePtr, 9) # 9 = Restaurar la ventana
