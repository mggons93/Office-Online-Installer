# Función para reiniciar el script con privilegios de administrador
function Start-ProcessAsAdmin {
    param (
        [string]$file,
        [string[]]$arguments = @()
    )
    Start-Process -FilePath $file -ArgumentList $arguments -Verb RunAs
}

# Comprobar si el script se está ejecutando como administrador
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    # Si no está ejecutándose como administrador, relanza el script con privilegios elevados
    Start-ProcessAsAdmin -file "powershell.exe" -arguments "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    exit
}


$ErrorActionPreference = "Stop"
# Enable TLSv1.2 for compatibility with older clients for current session
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

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

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft Office Online By Mggons Ver 2.2" Height="420" Width="450" Background="#778899">
    <Grid>
        <!-- Fila 1 -->
        <TextBlock Text="Seleccionar Office 365:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
        <ComboBox x:Name="variantComboBox" HorizontalAlignment="Left" VerticalAlignment="Top" Width="200" Margin="10,30,0,0">
            <!-- Office 365 -->
            <ComboBoxItem Content="Office 365 Pro Plus" Tag="O365ProPlusRetail"/>
            <ComboBoxItem Content="Office 365 Pro Plus (No Teams)" Tag="O365ProPlusEEANoTeamsRetail"/>
            <ComboBoxItem Content="Office 365 Business" Tag="O365BusinessRetail"/>
            <ComboBoxItem Content="Office 365 Business (No Teams)" Tag="O365BusinessEEANoTeamsRetail"/>
            <ComboBoxItem Content="Office 365 Education Cloud" Tag="O365EduCloudRetail"/>
            <ComboBoxItem Content="Office 365 Home Premium" Tag="O365HomePremRetail"/>
            <ComboBoxItem Content="Office 365 Small Business Premium" Tag="O365SmallBusPremRetail"/>
            <ComboBoxItem Content="Office 365 Project Standard" Tag="ProjectStdRetail"/>
            <ComboBoxItem Content="Office 365 Project Professional" Tag="ProjectProRetail"/>
            <ComboBoxItem Content="Office 365 Visio Standard" Tag="VisioStdRetail"/>
            <ComboBoxItem Content="Office 365 Visio Professional" Tag="VisioProRetail"/>
            <ComboBoxItem Content="---------------------"/>
            <!-- Edicion 2024 -->
            <ComboBoxItem Content="Office Professional Plus 2024" Tag="ProPlus2024Retail"/>
            <ComboBoxItem Content="Home &amp; Business 2024" Tag="HomeBusiness2024Retail"/>
            <ComboBoxItem Content="Office Home 2024" Tag="Home2024Retail"/>
            <ComboBoxItem Content="Project Standard 2024" Tag="ProjectStd2024Retail"/>
            <ComboBoxItem Content="Project Professional 2024" Tag="ProjectPro2024Retail"/>
            <ComboBoxItem Content="Visio Standard 2024" Tag="VisioStd2024Retail"/>
            <ComboBoxItem Content="Visio Professional 2024" Tag="VisioPro2024Retail"/>
            <ComboBoxItem Content="---------------------"/>
            <!-- Edición 2021 -->
            <ComboBoxItem Content="Office Professional Plus 2021" Tag="ProPlus2021Retail"/>
            <ComboBoxItem Content="Office Home &amp; Business 2021" Tag="HomeBusiness2021Retail"/>
            <ComboBoxItem Content="Office Home &amp; Student 2021" Tag="HomeStudent2021Retail"/>
            <ComboBoxItem Content="Office Standart 2021" Tag="Standard2021Retail"/>
            <ComboBoxItem Content="Project Standard 2021" Tag="ProjectStd2021Retail"/>
            <ComboBoxItem Content="Project Professional 2021" Tag="ProjectPro2021Retail"/>
            <ComboBoxItem Content="Visio Standard 2021" Tag="VisioStd2021Retail"/>
            <ComboBoxItem Content="Visio Professional 2021" Tag="VisioPro2021Retail"/>
            <ComboBoxItem Content="---------------------"/>
            <!-- Edición 2019 -->
            <ComboBoxItem Content="Office Professional Plus 2019" Tag="ProPlus2019Retail"/>
            <ComboBoxItem Content="Office Home &amp; Business 2019" Tag="HomeBusiness2019Retail"/>
            <ComboBoxItem Content="Office Home &amp; Student 2019" Tag="HomeStudent2019Retail"/>
            <ComboBoxItem Content="Office Standart 2019" Tag="Standard2019Retail"/>
            <ComboBoxItem Content="Project Standard 2019" Tag="ProjectStd2019Retail"/>
            <ComboBoxItem Content="Project Professional 2019" Tag="ProjectPro2019Retail"/>
            <ComboBoxItem Content="Visio Standard 2019" Tag="VisioStd2019Retail"/>
            <ComboBoxItem Content="Visio Professional 2019" Tag="VisioPro2019Retail"/>
            <ComboBoxItem Content="---------------------"/>
            <!-- Edición 2016 -->
            <ComboBoxItem Content="Office Professional Plus 2016" Tag="ProPlusRetail"/>
            <ComboBoxItem Content="Office Home &amp; Business 2016" Tag="HomeBusinessRetail"/>
            <ComboBoxItem Content="Office Home &amp; Student 2016" Tag="HomeStudentRetail"/>
            <ComboBoxItem Content="Office Standart 2016" Tag="StandardRetail"/>
            <ComboBoxItem Content="Project Standard 2016" Tag="ProjectStdRetail"/>
            <ComboBoxItem Content="Project Professional 2016" Tag="ProjectProRetail"/>
            <ComboBoxItem Content="Visio Standard 2016" Tag="VisioStdRetail"/>
            <ComboBoxItem Content="Visio Professional 2016" Tag="VisioProRetail"/>            
        </ComboBox>
        <CheckBox x:Name="architectureCheckBox" Content="Sistema operativo x64" IsChecked="$($is64Bit)" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="220,30,0,0" IsEnabled="True"/>

        <!-- Fila 2 -->
        <TextBlock Text="Seleccion de Idioma:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,70,0,0"/>
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
            <!-- Agrega aquí más opciones si es necesario -->
        </ComboBox>
        <CheckBox x:Name="vlActivationCheckBox" Content="Usar Activacion VL" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="220,90,0,0"/>
        
        <!-- Opciones de Activación VL solo si se activa VL -->
        <StackPanel x:Name="vlOptionsPanel" Visibility="Collapsed" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,130,0,0" Orientation="Horizontal">
            <StackPanel Orientation="Vertical" Margin="0,0,10,0">
                <TextBlock Text="Seleccione la edicion por Volumen:" HorizontalAlignment="Left" Margin="0,0,0,5"/>
                <ComboBox x:Name="editionComboBox" HorizontalAlignment="Left" Width="200">
		    <ComboBoxItem Content="****Office 2016****"/>
                    <ComboBoxItem Content="Office 2016 Standard VL"/>
                    <ComboBoxItem Content="Office 2016 Professional VL"/>
		    <ComboBoxItem Content="---------------------"/>
                    <ComboBoxItem Content="Project 2016 Standard VL"/>
                    <ComboBoxItem Content="Project 2016 Professional VL"/>
		    <ComboBoxItem Content="---------------------"/>
                    <ComboBoxItem Content="Visio 2016 Standard VL"/>
                    <ComboBoxItem Content="Visio 2016 Professional VL"/>
		    <ComboBoxItem Content="****Office 2019****"/>
                    <ComboBoxItem Content="Office 2019 Standard VL"/>
                    <ComboBoxItem Content="Office 2019 Professional VL"/>
		    <ComboBoxItem Content="---------------------"/>
                    <ComboBoxItem Content="Project 2019 Standard VL"/>
                    <ComboBoxItem Content="Project 2019 Professional VL"/>
		    <ComboBoxItem Content="---------------------"/>
                    <ComboBoxItem Content="Visio 2019 Standard VL"/>
                    <ComboBoxItem Content="Visio 2019 Professional VL"/>
		    <ComboBoxItem Content="****Office 2021****"/>
                    <ComboBoxItem Content="Office 2021 Standard VL"/>
                    <ComboBoxItem Content="Office 2021 Professional VL"/>
		    <ComboBoxItem Content="---------------------"/>
                    <ComboBoxItem Content="Project 2021 Standard VL"/>
                    <ComboBoxItem Content="Project 2021 Professional VL"/>		    
		    <ComboBoxItem Content="---------------------"/>
                    <ComboBoxItem Content="Visio 2021 Standard VL"/>
                    <ComboBoxItem Content="Visio 2021 Professional VL"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Orientation="Vertical">
                <TextBlock Text="Ingrese la clave de licencia:" HorizontalAlignment="Left" Margin="0,0,0,7"/>
                <TextBox x:Name="licenseKeyTextBox" HorizontalAlignment="Left" Width="200"/>
            </StackPanel>
        </StackPanel>

        <!-- Botón de instalación y log -->
        <Button x:Name="installButton" Content="Instalar" HorizontalAlignment="Left" VerticalAlignment="Top" Width="100" Height="30" Margin="10,200,0,0"/>
        <CheckBox x:Name="autoActivationCheckBox" Content="Activacion Automatica" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="120,205,0,0"/>
        <TextBlock Text="Log:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,240,0,0"/>
        <TextBox x:Name="logTextBox" HorizontalAlignment="Left" VerticalAlignment="Top" Width="410" Height="100" Margin="10,260,0,0" IsReadOnly="True" VerticalScrollBarVisibility="Auto"/>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$variantComboBox = $window.FindName("variantComboBox")
$languageComboBox = $window.FindName("languageComboBox")
$vlActivationCheckBox = $window.FindName("vlActivationCheckBox")
$architectureCheckBox = $window.FindName("architectureCheckBox")
$editionComboBox = $window.FindName("editionComboBox")
$licenseKeyTextBox = $window.FindName("licenseKeyTextBox")
$installButton = $window.FindName("installButton")
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
})

$autoActivationCheckBox.Add_Unchecked({
    if (-not $vlActivationCheckBox.IsChecked) {
        $autoActivationCheckBox.IsChecked = $false
    }
})

$installButton.Add_Click({
    $selectedVariant = $variantComboBox.SelectedItem.Tag
    $selectedLanguageFull = $languageComboBox.SelectedItem.Tag
    $selectedLanguage = $selectedLanguageFull -replace '.*\[(.*?)\]', '$1'
    $architecture = if ($architectureCheckBox.IsChecked) { "x64" } else { "x86" }
    $useVL = $vlActivationCheckBox.IsChecked
    $editionVL = if ($useVL) { $editionComboBox.SelectedItem.Content } else { $null }
    $licenseKey = if ($useVL) { $licenseKeyTextBox.Text } else { $null }
    $autoActivate = $autoActivationCheckBox.IsChecked
    
    $url = "https://%blank%c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=%blank%$selectedVariant&platform=$architecture%blank%&language=$selectedLanguage&version=O16GA"
    $url = $url -replace "%blank%", ""
    $outputFile = "$env:TEMP\${selectedVariant}-%blank%${selectedLanguage}-%blank%${architecture}.exe"
    $outputFile = $outputFile -replace "%blank%", ""
    
    $message = "Se procedera a descargar e instalar la siguiente variante de Office 365. ¿Desea continuar?"
    if ($useVL) {
        $message += "`n- Edición VL: $editionVL"
        $message += "`n- Clave de licencia: $licenseKey"
    }
    $message += "`n- Variante: $selectedVariant"
    $message += "`n- Idioma: $selectedLanguage"
    $message += "`n- Sistema: $architecture"
    $message += "`n- Activacion Automatica: $($autoActivate -eq $true)"
    $result = [System.Windows.MessageBox]::Show($message, "Confirmar Instalacion", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)

    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Add-LogMessage "Descargando $selectedVariant $selectedLanguage para sistema $architecture..."
		Add-LogMessage "Descarga completada. Iniciando la instalacion..."
        Invoke-WebRequest -Uri $url -OutFile $outputFile
       
        Start-Process -FilePath $outputFile -Wait
        "taskkill /f /im OfficeC2RClient.exe" | cmd

        Add-LogMessage "Instalacion completada."

        if ($autoActivate -and -not $useVL) {
            Add-LogMessage "Iniciando Activacion. Espere..."
            
            $url = "https://raw.%blank%githubuser%blank%content.com/mggo%blank%ns93/Mgg%blank%ons/main/Vali%blank%date/MAS%blank%_AIO.cmd"
            $url = $url -replace "%blank%", ""
	    $outputPath1 = "$env:TEMP\O%blank%hook_Acti%blank%vation_AI%blank%O.cmd"
     	    $outputPath1 = $outputPath1 -replace "%blank%", ""
            Add-LogMessage "Activando..."
            Invoke-WebRequest -Uri $url -OutFile $outputPath1
            Start-Process -FilePath $outputPath1 /Ohook -Wait 
            Add-LogMessage "Eliminando Archivos Usados..."
	        Remove-Item -Path $outputFile -Force
            Remove-Item -Path $outputPath1 -Force
            Add-LogMessage "Activacion completada."
        }

       if ($useVL) {
        Add-LogMessage "Convirtiendo Office Retail a Vol. Espere..."
        
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
        }

        Add-LogMessage "Edición seleccionada: $editionVL ($edition)"

        if (Test-Path "$env:ProgramFiles\Microsoft Office\Office16") {
            Set-Location "$env:ProgramFiles\Microsoft Office\Office16"
            Add-LogMessage "Cambiado al directorio: $env:ProgramFiles\Microsoft Office\Office16"
        } elseif (Test-Path "$env:ProgramFiles(x86)\Microsoft Office\Office16") {
            # Cambiar al directorio de Microsoft Office (32 bits)
            Set-Location "$env:ProgramFiles(x86)\Microsoft Office\Office16"
            Add-LogMessage "Cambiado al directorio: $env:ProgramFiles(x86)\Microsoft Office\Office16"
        } else {
            $errorMsg = "No se encontró la instalación de Microsoft Office en Office16."
            Add-LogMessage $errorMsg
            Write-Error $errorMsg
            exit
        }


        $licenseFiles = Get-ChildItem -Path "..\root\Licenses16\" -Filter "$edition*.xrm-ms"
        Add-LogMessage "Archivos de licencia encontrados: $($licenseFiles.Count)"
        
        # Iterar sobre cada archivo y ejecutar el comando cscript
        foreach ($file in $licenseFiles) {
            $filePath = "..\root\Licenses16\$($file.Name)"
            Add-LogMessage "Instalando Archivo: $filePath"
            cscript ospp.vbs /inslic:"$filePath"
        }
            Add-LogMessage "Instalando licencia: $licenseKey"
            cscript ospp.vbs /inpkey:$licenseKey
            cscript ospp.vbs /act
            Add-LogMessage "Activacion completada."
        }
    }
})

$window.ShowDialog() | Out-Null
[Win32]::ShowWindow($consolePtr, 9) # 9 = Restaurar la ventana
