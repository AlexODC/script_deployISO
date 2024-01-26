# Définition des couleurs
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
        foreach ($update in $updates) {
            Write-Host "Installation de la mise à jour : $($update.Title)" -ForegroundColor $GreenColor
            Install-WindowsUpdate -KBArticleID $update.KBArticleID -AutoReboot:$false -Confirm:$false
        }
        Write-Host "Toutes les mises à jour ont été installées. Un redémarrage pourrait être nécessaire." -ForegroundColor $OrangeText
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

# Fonction pour détecter la version d'Office installée
Function Detect-InstalledOffice {
    # Ajoutez ici votre logique pour détecter la version installée d'Office
}

# Fonction pour installer Office
Function Install-Office {
    # Votre code pour installer Office
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
        "1" {
            Install-WindowsUpdates
        }
        "2" {
            Write-Host "Entrez le nouveau nom de l'ordinateur :" -ForegroundColor $OrangeText
            $newName = Read-Host
            Rename-MyComputer -newName $newName
        }
        "3" {
            Update-WindowsStoreApps
        }
        "4" {
            Do {
                Write-Host "1. Installer Office"
                Write-Host "2. Installer OwnCloud"
                Write-Host "3. Installer Cybereason"
                Write-Host "4. Retour"

                Write-Host "Entrez votre choix :" -ForegroundColor $OrangeText -NoNewline
                $appChoice = Read-Host

                Switch ($appChoice) {
                    "1" { Install-Office }
                    "2" {
                        $downloadLink = "https://download.owncloud.com/desktop/ownCloud/stable/latest/win/ownCloud-latest.msi"
                        Download-And-Install-App -downloadLink $downloadLink -appName "ownCloud-latest.msi"
                    }
                    "3" {
                        $downloadLink = "https://link-to-cybereason-sensor/download"
                        Download-And-Install-App -downloadLink $downloadLink -appName "CybereasonSensor.exe"
                    }
                    "4" { break }
                }
            } While ($appChoice -ne "4")
        }
        "5" {
            Write-Host "Fin du script." -ForegroundColor $GreenColor
            break
        }
    }
} While ($selection -ne "5")
