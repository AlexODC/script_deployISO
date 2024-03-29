﻿# Définition des couleurs
$GreenColor = "Green"
$RedColor = "Red"
$OrangeText = "DarkYellow"

# Vérifier si le script est lancé en tant qu'administrateur
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Vérifier et installer le fournisseur NuGet si nécessaire
Function Check-And-Install-NuGet {
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet)) {
        Install-PackageProvider -Name NuGet -Force
        Import-PackageProvider -Name NuGet -Force
    }
}

# Fonction pour installer toutes les mises à jour Windows, y compris les facultatives
Function Install-WindowsUpdates {
    Check-And-Install-NuGet

    if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
        Remove-Module -Name PSWindowsUpdate -Force -ErrorAction SilentlyContinue
    }
    Install-Module -Name PSWindowsUpdate -Force -AllowClobber

    $updates = Get-WindowsUpdate -AcceptAll -IgnoreReboot
    if ($updates.Count -gt 0) {
        $updates | ForEach-Object {
            Write-Host $_.Title -ForegroundColor $GreenColor
        }
    } else {
        Write-Host "Aucune mise à jour Windows disponible." -ForegroundColor $GreenColor
    }
}

# Fonction pour renommer l'ordinateur
Function Rename-MyComputer {
    Param (
        [string]$newName
    )
    Rename-Computer -NewName $newName
    Write-Host "L'ordinateur a été renommé en '$newName'. Un redémarrage pourrait être nécessaire." -ForegroundColor $GreenColor
}

# Fonction pour mettre à jour les applications du Windows Store
Function Update-WindowsStoreApps {
    $wingetPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
    if (Test-Path $wingetPath) {
        winget upgrade --all | Out-Null
        Write-Host "Les mises à jour des applications Windows Store ont été effectuées." -ForegroundColor $GreenColor
    } else {
        Write-Host "Winget n'est pas installé." -ForegroundColor $RedColor
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

    # Mise à jour du fichier XML en fonction de la version choisie
    Switch ($officeChoice) {
        1 {
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

    Write-Host "Installation de la version choisie d'Office en cours..." -ForegroundColor $GreenColor
    Start-Process -FilePath "$ODTPath\setup.exe" -ArgumentList "/configure `"$XMLPath`"" -NoNewWindow -Wait
}

# Menu principal
Do {
    Write-Host "1. Installer les mises à jour Windows"
    Write-Host "2. Renommer cet ordinateur"
    Write-Host "3. Mettre à jour les applications du Windows Store"
    Write-Host "4. Installer des applications supplémentaires"
    Write-Host "5. Quitter"

    Write-Host "Entrez votre choix :" -ForegroundColor $OrangeText -NoNewline
    $selection = Read-Host

    Switch ($selection) {
        1 {
            Install-WindowsUpdates
        }
        2 {
            Write-Host "Entrez le nouveau nom de l'ordinateur :" -ForegroundColor $OrangeText
            $newName = Read-Host
            Rename-MyComputer -newName $newName
        }
        3 {
            Update-WindowsStoreApps
        }
        4 {
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
                        $downloadLink = "https://link-to-cybereason-sensor/download"
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
} While ($selection -ne '5')