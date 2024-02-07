Add-Type -AssemblyName PresentationFramework

$app_Dispo = @('1. Installer Office','2. Installer OwnCloud','3. Installer CyberReason')
$office_dispo = @('1.Office 2019 Standard', '2.Office 2019 Pro Plus', '3.Office 2021 Standard', '4.Office 2021 Pro Plus', '5.Office 365 Business', '6.Office SPLA')
$office_product_id = @('Standard2019Retail', 'ProPlus2019Retail', 'Standard2021Retail' ,'ProPlus2021Retail' ,'O365BusinessRetail','')
$add_to_domain = "False"

# Définition des couleurs
$GreenColor = "Green"
$RedColor = "Red"
$OrangeText = "DarkYellow"

# Vérifier si le script est lancé en tant qu'administrateur
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

#Déclaration de l'interface du menu principal
[xml]$XML_Menu = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="Modern Script Place" Height="380" Width="430">
    <Grid Margin="0,0,10,-6">
        <Button Name="btn_maj_pc" Content="Mettre à jour le PC (Windows et applications Windows Store)" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_rename_pc" Content="Renommer cet ordinateur et le mettre sur un domaine" HorizontalAlignment="Left" Margin="10,91,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_install_software" Content="Installer des applications supplémentaires" HorizontalAlignment="Left" Margin="10,172,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,253,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Label Content="@Copyright CHOCHOIS Alex et LECOUTRE Antoine les GOAT" HorizontalAlignment="Center" Margin="0,354,0,0" VerticalAlignment="Top"/>

    </Grid>
</Window>
"@

#Déclaration de l'interface de la mise à jour du poste
[xml]$XML_Maj_Poste = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
     Title="Windows_Maj" Height="470" Width="820">
    <Grid>
        <ListBox Name="List_Command" Margin="0,0,0,90"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_install_maj" Content="Installer les Maj" HorizontalAlignment="Left" Margin="400,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>

    </Grid>
</Window>
"@

#Déclaration de l'interface d'installation des applications
[xml]$XML_Install_App = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="Ultimate Installer Place" Height="470" Width="820">
    <Grid>
        <ListBox Name="List_Application" Margin="0,0,266,111"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_install" Content="Installer" HorizontalAlignment="Left" Margin="593,0,0,0" VerticalAlignment="Center" Height="76" Width="181"/>
    </Grid> 
</Window>
"@

#Déclaration de l'interface d'installation des offices
[xml]$XML_Install_Office = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="Ultimate Office Place" Height="470" Width="820">
    <Grid>
        <ListBox Name="List_Office" Margin="0,0,266,111"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_install" Content="Installer" HorizontalAlignment="Left" Margin="593,0,0,0" VerticalAlignment="Center" Height="76" Width="181"/>
    </Grid> 
</Window>
"@

#Déclaration de l'interface de renommage de poste
[xml]$XML_Rename_Poste = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="How to rename It (Place)" Height="470" Width="820">
    <Grid>
        <RadioButton GroupName="rb_Domaine" Name="rb_domaine_oui" Content="Oui" HorizontalAlignment="Left" Margin="10,33,0,0" VerticalAlignment="Top"/>
        <RadioButton GroupName="rb_Domaine" Name="rb_domaine_non" Content="Non" IsChecked="True" HorizontalAlignment="Left" Margin="10,53,0,0" VerticalAlignment="Top"/>
        <Label Content="Voulez-vous intégrer le PC à un domaine ?" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <Label Content="Nom du PC" HorizontalAlignment="Left" Margin="10,95,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txt_nom_pc" HorizontalAlignment="Left" Margin="10,121,0,0" TextWrapping="Wrap" Text="PC-" VerticalAlignment="Top" Width="234"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Label Content="Nom du Domaine" HorizontalAlignment="Left" Margin="400,10,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txt_domaine" HorizontalAlignment="Left" Margin="400,36,0,0" TextWrapping="Wrap" Text="Domaine.local" VerticalAlignment="Top" Width="234" IsEnabled="False" />
        <Button Name="btn_go" Content="Renommer le poste et/ou rejoindre le domaine" HorizontalAlignment="Left" Margin="10,267,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Label Content="Nom du compte de mise en domaine" HorizontalAlignment="Left" Margin="400,95,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txt_user" HorizontalAlignment="Left" Margin="400,121,0,0" TextWrapping="Wrap" Text="Administrateur" VerticalAlignment="Top" Width="234" IsEnabled="False" />
        <Button Name="btn_quitter_domaine" Content="Quitter le domaine" HorizontalAlignment="Left" Margin="400,267,0,0" VerticalAlignment="Top" Height="76" Width="390" IsEnabled="True"/>
    </Grid> 
</Window>
"@


##################################################################################################
#Setup de l'interface XML_Menu
$FormXML_Menu = (New-Object System.Xml.XmlNodeReader $XML_Menu)
$Window_Menu = [Windows.Markup.XamlReader]::Load($FormXML_Menu)

#Setup de l'interface XML_Maj_Poste
$FormXML_Maj_Poste = (New-Object System.Xml.XmlNodeReader $XML_Maj_Poste)
$Window_Maj_Poste = [Windows.Markup.XamlReader]::Load($FormXML_Maj_Poste)

#Setup de l'interface XML_Install_App
$FormXML_Install_App = (New-Object System.Xml.XmlNodeReader $XML_Install_App)
$Window_Install_App = [Windows.Markup.XamlReader]::Load($FormXML_Install_App)

#Setup de l'interface XML_Install_Office
$FormXML_Install_Office = (New-Object System.Xml.XmlNodeReader $XML_Install_Office)
$Window_Install_Office = [Windows.Markup.XamlReader]::Load($FormXML_Install_Office)

#Setup de l'interface XML_Rename_Poste
$FormXML_Rename_Poste = (New-Object System.Xml.XmlNodeReader $XML_Rename_Poste)
$Window_Rename_Poste = [Windows.Markup.XamlReader]::Load($FormXML_Rename_Poste)

##################################################################################################

#Déclaration des actions des boutons du menu
$Window_Menu.FindName("btn_quitter").add_click({ 
    $Window_Menu.Close() 
})

$Window_Menu.FindName("btn_maj_pc").add_click({
    #$Window_Maj_Poste.ShowDialog()
    Update-PC
})

$Window_Menu.FindName("btn_rename_pc").add_click({ 
    $Window_Rename_Poste.ShowDialog()
})

$Window_Menu.FindName("btn_install_software").add_click({ 
    $Window_Install_App.ShowDialog()
})


##################################################################################################

#Déclaration des actions des boutons du menu MAJ
$Window_Maj_Poste.FindName("btn_quitter").add_click({ 
    $Window_Maj_Poste.Hide()
})

$Window_Maj_Poste.FindName("btn_install_maj").add_click({ 
    Update-PC
})

##################################################################################################

#Déclaration des actions des boutons du menu Rename
$Window_Rename_Poste.FindName("btn_quitter").add_click({ 
    $Window_Rename_Poste.Hide()
})

$Window_Rename_Poste.FindName("btn_quitter_domaine").add_click({ 
    Remove-Computer -UnjoinDomainCredential Domain01\Admin01 -WorkgroupName "Local"
})

$Window_Rename_Poste.FindName("rb_domaine_non").add_click({ 
    $Window_Rename_Poste.FindName("txt_domaine").IsEnabled = $false
    $Window_Rename_Poste.FindName("txt_user").IsEnabled = $false
})

$Window_Rename_Poste.FindName("rb_domaine_oui").add_click({ 
    $Window_Rename_Poste.FindName("txt_domaine").IsEnabled = $true
    $Window_Rename_Poste.FindName("txt_user").IsEnabled = $true
})

$Window_Rename_Poste.FindName("btn_go").add_click({
    $newName = $Window_Rename_Poste.FindName("txt_nom_pc").Text
    $newDomaine = $Window_Rename_Poste.FindName("txt_domaine").Text
    $compteDomaine = $Window_Rename_Poste.FindName("txt_user").Text

    if(!$Window_Rename_Poste.FindName("txt_user").IsEnabled){
        Add-Computer -ComputerName $newName -DomainName $newDomaine -Credential $compteDomaine
    } else {
        Rename-Computer -NewName $newName
    }
})

##################################################################################################

#Déclaration des actions des boutons du menu Install app
$Window_Install_App.FindName("btn_quitter").add_click({ 
    $Window_Install_App.Hide()
})

$Window_Install_App.FindName("btn_install").add_click({
    $app_to_install = $Window_Install_App.FindName("List_Application").selectedItems
    $app_splited = $app_to_install.split(".")[0]
    Switch ($app_splited) {
        1 { 
            #Install-Office 
            $Window_Install_Office.ShowDialog()
        }
        2 {
            $downloadLink = "https://o360.odc.fr/s/b2RDcDbQAOw2bMQ/download"
            Download-And-Install-App -downloadLink $downloadLink -appName "ownCloud-5.2.1.13040.x64.msi"
        }
        3 {
            $downloadLink = "https://o360.odc.fr/s/k7OT8FI8UXYdeYG/download"
            Download-And-Install-App -downloadLink $downloadLink -appName "CybereasonSensor.exe"
        }
    }
})

##################################################################################################

#Déclaration des actions des boutons du menu Install Office
$Window_Install_Office.FindName("btn_quitter").add_click({ 
    $Window_Install_Office.Hide()
})

$Window_Install_Office.FindName("btn_install").add_click({
    Write-Host "Lancement de l'installation d'office"
    $office_to_install = $Window_Install_Office.FindName("List_Office").selectedItems
    $office_splited = $office_to_install.split(".")

    $officeVersion = $office_splited[1]
    $XMLContent = 
@"
    <Configuration>
        <Add OfficeClientEdition="64" Channel="Monthly">
            <Product ID="$office_product_id[$office_splited[0]-1]">
                <Language ID="fr-fr" />
            </Product>
        </Add>
        <Display Level="None" AcceptEULA="TRUE" />
        <Property Name="AUTOACTIVATE" Value="1"/>
    </Configuration>
"@
    if($office_product_id[$office_splited[0]] = 6){
        $officeVersion = "Office SPLA"
        $SPLADownloadLink = "https://o360.odc.fr/s/QfjXvKHGUnUu2Xj/download"
        $SPLAExePath = "$env:TEMP\Install_Office_2021_SPLA.exe"
            
        # Téléchargement du fichier d'installation SPLA
        Write-Host "Téléchargement de l'installation pour Office SPLA..."
        Invoke-WebRequest -Uri $SPLADownloadLink -OutFile $SPLAExePath
            
        # Exécution du fichier d'installation SPLA
        Write-Host "Installation de Office SPLA en cours..."
        Start-Process -FilePath $SPLAExePath -ArgumentList "/silent" -NoNewWindow -Wait
    } else {
        Install-Office
    }
})

##################################################################################################

# Vérifier et installer le fournisseur NuGet si nécessaire
Function Check-And-Install-NuGet {
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet)) {
        Install-PackageProvider -Name NuGet -Force
        Import-PackageProvider -Name NuGet -Force
    }
}

# Fonction pour télécharger et installer une application depuis un lien public OwnCloud, Cybereason, etc.
Function Download-And-Install-App {
    Param (
        [string]$downloadLink,
        [string]$appName
    )
    $downloadsPath = [System.Environment]::GetFolderPath('UserProfile') + '\Downloads'
    $localPath = Join-Path -Path $downloadsPath -ChildPath $appName

    try {
        Invoke-WebRequest -Uri $downloadLink -OutFile $localPath
        if (Test-Path $localPath) {
            if ($appName -like "*.msi") {
                Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$localPath`" /qn" -Wait
                Write-Host "$appName a été installée avec succÃ¨s." -ForegroundColor $GreenColor
            } else {
                Start-Process -FilePath $localPath -Args "/S" -Wait
                Write-Host "$appName a été installée avec succÃ¨s." -ForegroundColor $GreenColor
            }
        } else {
            Write-Host "Erreur : Impossible de trouver le fichier téléchargé." -ForegroundColor $RedColor
        }
    } catch {
        Write-Host "Erreur lors du téléchargement ou de l'installation : $_" -ForegroundColor $RedColor
    }
    Remove-Item -Path $localPath -Force
}

# Fonction pour détecter la version d'Office installée
Function Detect-InstalledOffice {
    $Folder = 'C:\Program Files\Microsoft Office 15'
    if (Test-Path -Path $Folder) {
        Write-Host "Un office est déja présent !"
        return true
    } else {
        return false
    }
}

# Fonction pour installer Office
Function Install-Office {
    $ODTDownloadLink = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17126-20132.exe"
    $ODTPath = "$env:TEMP\ODT"
    $XMLPath = "$ODTPath\configuration.xml"

    # Téléchargement et extraction de l'Office Deployment Tool
    Invoke-WebRequest -Uri $ODTDownloadLink -OutFile "$env:TEMP\ODT.exe"
    Start-Process -FilePath "$env:TEMP\ODT.exe" -ArgumentList "/extract:`"$ODTPath`"" -NoNewWindow -Wait

    <#
    Write-Host "Choisissez la version d'Office à installer :"
    Write-Host "1. Office 2019 Standard"
    Write-Host "2. Office 2019 Pro Plus"
    Write-Host "3. Office 2021 Standard"
    Write-Host "4. Office 2021 Pro Plus"
    Write-Host "5. Office 365 Business"
    $officeChoice = Read-Host
    $officeVersion = ""

    # Mise à jour du fichier XML en fonction de la version choisie
    Switch ($officeChoice) {
        1 {
            $officeVersion = "Office 2019 Standard"
            $XMLContent = @"
            <Configuration>
                <Add OfficeClientEdition="64" Channel="Monthly">
                    <Product ID="Standard2019Retail">
                        <Language ID="fr-fr" />
                    </Product>
                </Add>
                <Display Level="None" AcceptEULA="TRUE" />
                <Property Name="AUTOACTIVATE" Value="1"/>
            </Configuration>
"@
        }
        2 {
            $officeVersion = "Office 2019 Pro Plus"
            $XMLContent = @"
            <Configuration>
                <Add OfficeClientEdition="64" Channel="Monthly">
                    <Product ID="ProPlus2019Retail">
                        <Language ID="fr-fr" />
                    </Product>
                </Add>
                <Display Level="None" AcceptEULA="TRUE" />
                <Property Name="AUTOACTIVATE" Value="1"/>
            </Configuration>
"@
        }
        3 {
            $officeVersion = "Office 2021 Standard"
            $XMLContent = @"
            <Configuration>
                <Add OfficeClientEdition="64" Channel="Monthly">
                    <Product ID="Standard2021Retail">
                        <Language ID="fr-fr" />
                    </Product>
                </Add>
                <Display Level="None" AcceptEULA="TRUE" />
                <Property Name="AUTOACTIVATE" Value="1"/>
            </Configuration>
"@
        }
        4 {
            $officeVersion = "Office 2021 Pro Plus"
            $XMLContent = @"
            <Configuration>
                <Add OfficeClientEdition="64" Channel="Monthly">
                    <Product ID="ProPlus2021Retail">
                        <Language ID="fr-fr" />
                    </Product>
                </Add>
                <Display Level="None" AcceptEULA="TRUE" />
                <Property Name="AUTOACTIVATE" Value="1"/>
            </Configuration>
"@
        }
        5 {
            $officeVersion = "Office 365 Business"
            $XMLContent = @"
            <Configuration>
                <Add OfficeClientEdition="64" Channel="Monthly">
                    <Product ID="O365BusinessRetail">
                        <Language ID="fr-fr" />
                    </Product>
                </Add>
                <Display Level="None" AcceptEULA="TRUE" />
                <Property Name="AUTOACTIVATE" Value="1"/>
            </Configuration>
"@
        }
    }
    #>
    $XMLContent | Out-File -FilePath $XMLPath

    # Détecter et désinstaller la version actuelle d'Office si présente
    $installedOffice = Detect-InstalledOffice
    if ($installedOffice) {
        Write-Host "Désinstallation de la version actuelle d'Office..."
        Remove-PreviousOfficeInstalls -Force -Quiet
    }

    # Installation de la nouvelle version d'Office
    Write-Host "Installation de $officeVersion en cours..." -ForegroundColor $GreenColor
    Start-Process -FilePath "$ODTPath\setup.exe" -ArgumentList "/configure `"$XMLPath`"" -NoNewWindow -Wait
}

# Fonction pour installer toutes les mises à jour Windows, y compris les facultatives
Function Install-WindowsUpdates {
    Write-Host "Lancement des mises à jours Windows"
    Check-And-Install-NuGet

    if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
        Remove-Module -Name PSWindowsUpdate -Force -ErrorAction SilentlyContinue
    }
    Install-Module -Name PSWindowsUpdate -Force -AllowClobber

    $updates = Get-WindowsUpdate -AcceptAll -IgnoreReboot
    $rebootRequired = $false

    if ($updates.Count -gt 0) {
        foreach ($update in $updates) {
            Write-Host "Installation de la mise à jour : $($update.Title)" -ForegroundColor $GreenColor
            Install-WindowsUpdate -KBArticleID $update.KBArticleID -AutoReboot:$false -Confirm:$false -IgnoreRebootRequired
            if ($update.IsRebootRequired) {
                $rebootRequired = $true
            }
        }
        if ($rebootRequired) {
            Write-Host "Redémarrage nécessaire pour terminer l'installation des mises à jour. Veuillez redémarrer votre ordinateur." -ForegroundColor $OrangeText
        } else {
            Write-Host "Toutes les mises à jour ont été installallées. Un redémarrage pourrait être nécessaire pour appliquer les mises à jour." -ForegroundColor $OrangeText
        }
    } else {
        Write-Host "Aucune mise à jour Windows disponible." -ForegroundColor $GreenColor
    }

    # Retour au menu principal ou fin de la fonction
    return
}

# Fonction pour mettre à jour les applications du Windows Store
Function Update-WindowsStoreApps {
    Write-Host "Lancement des mises à jours des applications"
    $wingetPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
    if (Test-Path $wingetPath) {
        winget upgrade --all --force --accept-package-agreements --accept-source-agreements
        Write-Host "Les mises à jour des applications Windows Store ont été effectuées." -ForegroundColor $GreenColor
    } else {
        Write-Host "Winget n'est pas installé." -ForegroundColor $RedColor
    }
    # Retour au menu principal ou fin de la fonction
    return 
}

# Fonction pour mettre à jour le PC (Windows et applications Windows Store)
Function Update-PC {
    Install-WindowsUpdates
    Update-WindowsStoreApps
}

# Les autres fonctions restent identiques...


##################################################################################################
foreach ($office in $office_dispo) {
    $Window_Install_Office.FindName("List_Office").Items.Add("$office")
}
foreach ($app in $app_Dispo) {
    $Window_Install_App.FindName("List_Application").Items.Add("$app")
}

cls

#Affichage de la interface
$Window_Menu.ShowDialog()

##################################################################################################
