$url = 'https://raw.githubusercontent.com/AlexODC/script_deployISO/main/Modern_Script_Place.ps1' # URL du script de déploiement
$desktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
$output = Join-Path $desktopPath "Modern_Script_Place.ps1"

# Télécharger le contenu du script en mémoire
$response = Invoke-WebRequest -Uri $url

# Vérifier si le contenu téléchargé n'est pas vide avant de l'écrire dans un fichier
if ($response.Content.Length -gt 0) {
    # Écrire le contenu en spécifiant l'encodage UTF-8
    $response.Content | Out-File -FilePath $output -Encoding UTF8

    Start-Process PowerShell -ArgumentList " -ExecutionPolicy Bypass -File `"$output`""

    # Supprimer le script de téléchargement seulement si le fichier téléchargé n'est pas vide
    $cheminDuFichier = Join-Path $desktopPath "Downloader_Modern_Script_Place.ps1"
    Remove-Item -Path $cheminDuFichier
} else {
    # Si le contenu est vide, afficher un message d'erreur ou prendre une autre action
    Write-Host "Le téléchargement du script a échoué ou le script est vide."
}
