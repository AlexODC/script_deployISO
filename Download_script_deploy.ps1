$url = "https://raw.githubusercontent.com/AlexODC/script_deployISO/main/Script_Deploy_d%C3%A9ploy%C3%A9.ps1" # URL du script de déploiement
$desktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
$output = Join-Path $desktopPath "script_deployISO.ps1"

# Télécharger le contenu du script en mémoire
$response = Invoke-WebRequest -Uri $url

# Écrire le contenu en spécifiant l'encodage UTF-8
$response.Content | Out-File -FilePath $output -Encoding UTF8

# Exécuter le script téléchargé
Start-Process PowerShell -ArgumentList "-ExecutionPolicy Bypass -File `"$output`""

# Supprimer le fichier téléchargé (Si nécessaire. Notez que la ligne actuelle semble supprimer un autre script. Assurez-vous que le chemin est correct.)
$cheminDuFichier = Join-Path $desktopPath "Download_script_deploy.ps1"
Remove-Item -Path $cheminDuFichier -Force
