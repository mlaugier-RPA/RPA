<#
# API Insee 
# Version : v1.0
# Maxime LAUGIER
# Update du 15/10/2025
#>

# === CONFIGURATION ===
$apiToken = "5fb9b51f-515c-4848-b9b5-1f515c184895"   # <-- ta clé API INSEE
$inputFile = "C:\Users\MaximeLaugier\Downloads\ListeClients_20251010 (1) - Copie.csv"
$outputFile = "C:\Users\MaximeLaugier\Downloads\resultats_sirets.csv"

$pauseMin = 700    # pause min entre requêtes (ms)
$pauseMax = 1500   # pause max entre requêtes (ms)
$maxRetry = 3      # nombre max de retries sur erreur 429

# === FONCTION POUR REQUÊTER L'API SIRENE POUR LES SIRET ===
function Get-SiretsFromSiren {
    param([string]$Siren)

    $url = "https://api.insee.fr/api-sirene/3.11/siret?q=siren:$Siren"
    $headers = @{
        "accept" = "application/json"
        "X-INSEE-Api-Key-Integration" = $apiToken
    }

    $retry = 0
    while ($retry -lt $maxRetry) {
        try {
            $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop
            $etablissements = $response.etablissements

            if (-not $etablissements) {
                Write-Warning "Aucun établissement trouvé pour SIREN $Siren"
                return $null
            }

            return $etablissements | ForEach-Object {
                [PSCustomObject]@{
                    SIREN = $Siren
                    SIRET = $_.siret
                    Etat = $_.etatAdministratifEtablissement
                    RaisonSociale = $_.uniteLegale.denominationUniteLegale
                    Adresse = (
                        "$($_.adresseEtablissement.numeroVoieEtablissement) " +
                        "$($_.adresseEtablissement.typeVoieEtablissement) " +
                        "$($_.adresseEtablissement.libelleVoieEtablissement) " +
                        "$($_.adresseEtablissement.codePostalEtablissement) " +
                        "$($_.adresseEtablissement.libelleCommuneEtablissement)"
                    ).Trim()
                }
            }
        }
        catch {
            if ($_.Exception.Message -match "429") {
                $retry++
                $waitTime = 2000 * $retry   # pause progressive: 2s, 4s, 6s
                Write-Warning "429 Too Many Requests pour SIREN $Siren. Pause $waitTime ms (retry $retry/$maxRetry)..."
                Start-Sleep -Milliseconds $waitTime
            }
            else {
                Write-Warning "Erreur API pour SIREN $Siren : $($_.Exception.Message)"
                return $null
            }
        }
    }

    Write-Warning "Échec après $maxRetry retries pour SIREN $Siren"
    return $null
}

# === TRAITEMENT DU CSV ===
$entreprises = Import-Csv -Path $inputFile -Delimiter ";" -Encoding UTF8
$resultats = @()

foreach ($entreprise in $entreprises) {
    $siren = $entreprise.Siren

    if ([string]::IsNullOrWhiteSpace($siren)) {
        Write-Host "⏭️  Ligne ignorée : SIREN vide."
        continue
    }

    $siren = $siren.Trim()

    if ($siren -match '^\d{9}$') {
        Write-Host "🔍 Recherche des SIRET pour SIREN $siren..."
        $sirets = Get-SiretsFromSiren -Siren $siren

        if ($sirets) {
            foreach ($siret in $sirets) {
                $ligne = $entreprise | Select-Object *
                $ligne | Add-Member -NotePropertyName "SiretTrouvé" -NotePropertyValue $siret.SIRET
                $ligne | Add-Member -NotePropertyName "EtatEtablissement" -NotePropertyValue $siret.Etat
                $ligne | Add-Member -NotePropertyName "RaisonSocialeAPI" -NotePropertyValue $siret.RaisonSociale
                $ligne | Add-Member -NotePropertyName "AdresseAPI" -NotePropertyValue $siret.Adresse
                $resultats += $ligne
            }
        } else {
            $ligne = $entreprise | Select-Object *
            $ligne | Add-Member -NotePropertyName "SiretTrouvé" -NotePropertyValue ""
            $ligne | Add-Member -NotePropertyName "EtatEtablissement" -NotePropertyValue "Introuvable"
            $ligne | Add-Member -NotePropertyName "RaisonSocialeAPI" -NotePropertyValue ""
            $ligne | Add-Member -NotePropertyName "AdresseAPI" -NotePropertyValue ""
            $resultats += $ligne
        }

        # Pause aléatoire entre chaque requête pour éviter le rate-limit
        $randomPause = Get-Random -Minimum $pauseMin -Maximum $pauseMax
        Start-Sleep -Milliseconds $randomPause
    } else {
        Write-Warning "SIREN invalide ou mal formaté : $siren"
    }
}

# === EXPORT DES RÉSULTATS ===
$resultats | Export-Csv -Path $outputFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
Write-Host "✅ Terminé : résultats exportés dans $outputFile"
