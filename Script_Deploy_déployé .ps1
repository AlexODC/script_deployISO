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
    $rebootRequired = $false

    if ($updates.Count -gt 0) {
        foreach ($update in $updates) {
            Write-Host "Installation de la mise à jour : $($update.Title)" -ForegroundColor $GreenColor
            Install-WindowsUpdate -KBArticleID $update.KBArticleID -AutoReboot:$false -Confirm:$false
            if ($update.IsRebootRequired) {
                $rebootRequired = $true
            }
        }
        if ($rebootRequired) {
            Write-Host "Redémarrage nécessaire pour terminer l'installation des mises à jour. Veuillez redémarrer votre ordinateur." -ForegroundColor $OrangeText
        } else {
            Write-Host "Toutes les mises à jour ont été installées. Un redémarrage pourrait être nécessaire." -ForegroundColor $OrangeText
        }
    } else {
        Write-Host "Aucune mise à jour Windows disponible." -ForegroundColor $GreenColor
    }

    # Retour au menu principal ou fin de la fonction
    return
}

# Fonction pour mettre à jour les applications du Windows Store
Function Update-WindowsStoreApps {
    $wingetPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
    #$wingetPath = "C:\Users\ODC\AppData\Local\Microsoft\WindowsApps\winget.exe"
    if (Test-Path $wingetPath) {
        winget upgrade --all --force
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
    Start-Sleep -Seconds 10
    cls
}

# Les autres fonctions restent identiques...

# Menu principal
Do {
    Write-Host "1. Mettre à jour le PC (Windows et applications Windows Store)"
    Write-Host "2. Renommer cet ordinateur"
    Write-Host "3. Installer des applications supplémentaires"
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
            Rename-MyComputer -newName $newName
        }
        "3" {
            Do {
                # Le sous-menu pour installer des applications supplémentaires...
            } While ($appChoice -ne "4")
        }
        "4" {
            Write-Host "Fin du script." -ForegroundColor $GreenColor
            break
        }
    }
} While ($selection -ne "4")
