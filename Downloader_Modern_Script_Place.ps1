$url = 'https://raw.githubusercontent.com/AlexODC/script_deployISO/main/Modern_Script_Place.ps1' # URL du script de déploiement
$desktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
$output = Join-Path $desktopPath "Modern_Script_Place.ps1"

# Télécharger le contenu du script en mémoire
$response = Invoke-WebRequest -Uri $url

# Écrire le contenu en spécifiant l'encodage UTF-8
$response.Content | Out-File -FilePath $output -Encoding UTF8

Start-Process PowerShell -ArgumentList " -ExecutionPolicy Bypass -File `"$output`""

$cheminDuFichier = Join-Path $desktopPath "Downloader_Modern_Script_Place.ps1"
Remove-Item -Path $cheminDuFichier
