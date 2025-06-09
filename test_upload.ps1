# Configuration
$ragicBaseUrl = "https://ragic.elyse.energy/default/"
$ragicApiKey = "Njl3OENtYnFnTExxSzNWVXZ6Y2E1Tlg0RWtjcVVBdnFHeVR0cTRCS09OWDMwZHlqRVc3WGx3WFJTNTFXMDRDZlZ2OWdXVElUaEtnPQ=="
$boundary = "---------------------------" + (Get-Random)
$testFilePath = "test.txt"

# Créer un fichier test
"Ceci est un fichier test pour l'upload Ragic" | Out-File -FilePath $testFilePath -Encoding utf8

Write-Host "[INFO] Preparation de la requete..."

# Préparer les données du formulaire
$formData = @{
    "1001623" = "Real"                    # Fake ? (Real/Fake)
    "1001060" = "Test Upload PowerShell"  # Name
    "1004024" = "Level 1"                 # Validator - level 1
    "1004025" = "Level 2"                 # Validator - level 2
    "1001044" = "1.0"                     # Version
    "1001068" = "Julien Fernandez"       # Author
    "1001069" = (Get-Date -Format "yyyy-MM-dd") # Delivery date
    "1001045" = "Test upload PowerShell"  # Change log
    "1001063" = "Internal simulation only (expert)" # Can be use for ?
    "1001040" = "test.txt"               # File name
    "1005174" = "Planning"               # Type
    "1001066" = "methanol"               # Main molecule/expertise
    "1001067" = "average per year"       # Main timescale
}

# Construire le corps de la requête
$LF = "`r`n"
$bodyLines = New-Object System.Collections.ArrayList

# Ajouter les champs de données
Write-Host "[INFO] Construction des champs de donnees..."
foreach ($key in $formData.Keys) {
    $null = $bodyLines.Add("--$boundary")
    $null = $bodyLines.Add("Content-Disposition: form-data; name=`"$key`"")
    $null = $bodyLines.Add("")
    $null = $bodyLines.Add($formData[$key])
}

# Ajouter le fichier
Write-Host "[INFO] Ajout du fichier..."
$fileContent = [System.IO.File]::ReadAllBytes($testFilePath)
$null = $bodyLines.Add("--$boundary")
$null = $bodyLines.Add("Content-Disposition: form-data; name=`"1001040`"; filename=`"$testFilePath`"")
$null = $bodyLines.Add("Content-Type: application/octet-stream")
$null = $bodyLines.Add("")

# Convertir les lignes en bytes
$bodyStart = [System.Text.Encoding]::UTF8.GetBytes(($bodyLines -join $LF) + $LF)
$bodyEnd = [System.Text.Encoding]::UTF8.GetBytes($LF + "--$boundary--" + $LF)

Write-Host "[INFO] Preparation de la requete HTTP..."
$url = $ragicBaseUrl + "simulation-files/1"

try {
    # Créer la requête
    $request = [System.Net.HttpWebRequest]::Create($url)
    $request.Method = "POST"
    $request.ContentType = "multipart/form-data; boundary=$boundary"
    
    # Ajouter l'authentification
    $request.Headers.Add("Authorization", "Basic " + $ragicApiKey)

    Write-Host "[INFO] Envoi des donnees..."
    Write-Host "[DEBUG] URL: $url"
    Write-Host "[DEBUG] Authorization: Basic " + $ragicApiKey
    Write-Host "[DEBUG] Content-Type: " + $request.ContentType
    Write-Host "[DEBUG] Boundary: $boundary"
    Write-Host "[DEBUG] Form data:"
    foreach ($key in $formData.Keys) {
        Write-Host "  $key = $($formData[$key])"
    }
    
    # Écrire le corps de la requête
    $requestStream = $request.GetRequestStream()
    $requestStream.Write($bodyStart, 0, $bodyStart.Length)
    $requestStream.Write($fileContent, 0, $fileContent.Length)
    $requestStream.Write($bodyEnd, 0, $bodyEnd.Length)
    $requestStream.Close()

    Write-Host "[INFO] Attente de la reponse..."
    # Obtenir la réponse
    $response = $request.GetResponse()
    $responseStream = $response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($responseStream)
    $responseContent = $reader.ReadToEnd()

    Write-Host "[SUCCESS] Reponse recue :"
    Write-Host $responseContent

    $reader.Close()
    $response.Close()
}
catch {
    Write-Host "[ERROR] Erreur lors de l'upload :"
    Write-Host $_.Exception.Message
    Write-Host "Response Status Code:" $_.Exception.Response.StatusCode.value__
    Write-Host "Response Status Description:" $_.Exception.Response.StatusDescription
    
    # Essayer de lire le corps de l'erreur
    try {
        $errorStream = $_.Exception.Response.GetResponseStream()
        $errorReader = New-Object System.IO.StreamReader($errorStream)
        $errorContent = $errorReader.ReadToEnd()
        Write-Host "Response Body:" $errorContent
    }
    catch {
        Write-Host "Impossible de lire le corps de l'erreur"
    }
}
finally {
    # Nettoyer
    Remove-Item $testFilePath -ErrorAction SilentlyContinue
    Write-Host "[INFO] Nettoyage effectue"
} 