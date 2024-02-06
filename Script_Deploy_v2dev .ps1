# DÃ©finition des couleurs
$GreenColor = "Green"
$RedColor = "Red"
$OrangeText = "DarkYellow"

# VÃ©rifier si le script est lancÃ© en tant qu'administrateur
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# VÃ©rifier et installer le fournisseur NuGet si nÃ©cessaire
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
                Write-Host "$appName a été installé avec succès." -ForegroundColor $GreenColor
            } else {
                Start-Process -FilePath $localPath -Args "/S" -Wait
                Write-Host "$appName a été installé avec succès." -ForegroundColor $GreenColor
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
    $officeVersion = Get-OfficeVersion -ShowAllInstalledProducts
    if ($officeVersion) {
        Write-Host "Version d'Office actuellement installée : $($officeVersion.DisplayName)" -ForegroundColor $GreenColor
        return $officeVersion
    } else {
        Write-Host "Aucune version d'Office n'est installée." -ForegroundColor $GreenColor
        return $null
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


# Fonction pour installer toutes les mises Ã  jour Windows, y compris les facultatives
Function Install-WindowsUpdates {
    Check-And-Install-NuGet

    if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
        Remove-Module -Name PSWindowsUpdate -Force -ErrorAction SilentlyContinue
    }
    Install-Module -Name PSWindowsUpdate -Force -AllowClobber

    $updates = Get-WindowsUpdate -AcceptAll -IgnoreReboot
    $rebootRequired = $false

    if ($updates.Count -gt 0) {
        foreach ($update in $updates) {
            Write-Host "Installation de la mise Ã  jour : $($update.Title)" -ForegroundColor $GreenColor
            Install-WindowsUpdate -KBArticleID $update.KBArticleID -AutoReboot:$false -Confirm:$false
            if ($update.IsRebootRequired) {
                $rebootRequired = $true
            }
        }
        if ($rebootRequired) {
            Write-Host "RedÃ©marrage nÃ©cessaire pour terminer l'installation des mises Ã  jour. Veuillez redÃ©marrer votre ordinateur." -ForegroundColor $OrangeText
        } else {
            Write-Host "Toutes les mises Ã  jour ont Ã©tÃ© installÃ©es. Un redÃ©marrage pourrait Ãªtre nÃ©cessaire." -ForegroundColor $OrangeText
        }
    } else {
        Write-Host "Aucune mise Ã  jour Windows disponible." -ForegroundColor $GreenColor
    }

    # Retour au menu principal ou fin de la fonction
    return
}

# Fonction pour mettre Ã  jour les applications du Windows Store
Function Update-WindowsStoreApps {
    $wingetPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
    #$wingetPath = "C:\Users\ODC\AppData\Local\Microsoft\WindowsApps\winget.exe"
    if (Test-Path $wingetPath) {
        winget upgrade --all --force
        Write-Host "Les mises Ã  jour des applications Windows Store ont Ã©tÃ© effectuÃ©es." -ForegroundColor $GreenColor
    } else {
        Write-Host "Winget n'est pas installÃ©." -ForegroundColor $RedColor
    }
    # Retour au menu principal ou fin de la fonction
    return 
}

# Fonction pour mettre Ã  jour le PC (Windows et applications Windows Store)
Function Update-PC {
    Install-WindowsUpdates
    Update-WindowsStoreApps
}

# Les autres fonctions restent identiques...

# Menu principal
Do {
    Write-Host "1. Mettre Ã  jour le PC (Windows et applications Windows Store)"
    Write-Host "2. Renommer cet ordinateur"
    Write-Host "3. Installer des applications supplÃ©mentaires"
    Write-Host "4. Quitter"

    Write-Host "Entrez votre choix :" -ForegroundColor $OrangeText -NoNewline
    $selection = Read-Host

    Switch ($selection) {
        "1" {
            Update-PC
        }
        "2" {
            Write-Host "Entrez le nouveau nom de l'ordinateur :" -ForegroundColor $OrangeText
            $newName = Read-Host
            Rename-Computer -newName $newName
        }
        "3" {
             Do {
                Write-Host "1. Installer Office"
                Write-Host "2. Installer OwnCloud"
                Write-Host "3. Installer Cybereason"
                Write-Host "4. Retour"

                Write-Host "Entrez votre choix :" -ForegroundColor $OrangeText -NoNewline
                $appChoice = Read-Host

                Switch ($appChoice) {
                    1 { Install-Office }
                    2 {
                        $downloadLink = "https://download.owncloud.com/desktop/ownCloud/stable/latest/win/ownCloud-latest.msi"
                        Download-And-Install-App -downloadLink $downloadLink -appName "ownCloud-latest.msi"
                    }
                    3 {
                        $downloadLink = "https://o360.odc.fr/s/k7OT8FI8UXYdeYG/download"
                        Download-And-Install-App -downloadLink $downloadLink -appName "CybereasonSensor.exe"
                    }
                    4 { break }
                }
            } While ($appChoice -ne '4')
        }
        5 {
            Write-Host "Fin du script." -ForegroundColor $GreenColor
            break
        }
    }
    Read-Host -Prompt "Appuyez sur ENTRER pour continuer." | Out-Null
    cls
} While ($selection -ne "4")
