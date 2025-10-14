<#
# API Orchestrator <=> PowerBI - Script Corrigé
# Version : v9.4
# Correction principale: Ajout d'un Start-Sleep pour éviter l'erreur HTTP 429 (Too Many Requests).
# Maxime LAUGIER
# Update du 14/10/2025
#>

# === variables fixes pour tout le script ===
$Org = "extiavqvkelj"
$Tenant = "DefaultTenant"
$XlsPath = "H:\Mon Drive\Reporting PowerBI\UiPathJobs.xlsx"
$BaseUrl = "https://cloud.uipath.com/$Org/$Tenant/orchestrator_/odata"


# === Paramètres ROI pour le calcul du gain ===
$CostPerHour = 30         # € / heure d’un humain
$MinutesSavedPerJob = 20  # minutes économisées par job réussi
$MonthlyRPACost = 5900    # € coût global RPA mensuel


# === authentification sur l'api et on essaie de récupérer le token ===
Write-Host "🔑 Tentative de récupération du jeton d'accès..."
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")

$body = "client_id=92615cee-13a8-4195-b52a-3543976033cc&client_secret=lOa%5EtVshMA!mLwLsI8kbwNO)8QH%23p1c%23Qa_jmIN%3FCkYo~YOevEs73EVc(Cb(N2jy&grant_type=client_credentials"

try {
    $response = Invoke-RestMethod "https://cloud.uipath.com/$Org/identity_/connect/token" -Method 'POST' -Headers $headers -Body $body
    $PAT = $response.access_token
} catch {
    Write-Host "❌ Erreur d'authentification ou réseau : $($_.Exception.Message)" -ForegroundColor Red
    exit
}

if (-not $PAT) { Write-Host "❌ Token introuvable" -ForegroundColor Red; exit }
Write-Host "✅ Jeton récupéré."


# === Après avoir récupérer le token dynamique, on set-up les headers pour le call API ===
$Headers = @{
    "Authorization" = "Bearer $PAT"
    "Accept" = "application/json;odata=nometadata"
}


# === Récupération des folders ===
try {
    $Folders = (Invoke-RestMethod -Uri "$BaseUrl/Folders" -Headers $Headers).value
} catch {
    Write-Host "❌ Erreur lors de la récupération des folders : $($_.Exception.Message)" -ForegroundColor Red
    exit
}
if (-not $Folders) { Write-Host "❌ Aucun folder trouvé" -ForegroundColor Yellow; exit }
Write-Host "📁 $(@($Folders).Count) folders trouvés."


# === Fonction pour avoir l'ID des folder d'UIpath Orchestrator et faire un for each pour checker chaque dossier ===
function Get-UipathJobsForFolder {
    param (
        [string]$FolderId,
        [string]$FolderName,
        [datetime]$StartDate
    )

    $FolderHeaders = @{ }
    foreach ($key in $Headers.Keys) { $FolderHeaders[$key] = $Headers[$key] }
    $FolderHeaders["X-UIPATH-OrganizationUnitId"] = "$FolderId"

    $Jobs = @()
    $FilterDate = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    
    # Le filtre CreationTime est utilisé pour inclure TOUS les jobs (terminés et en cours).
    $NextUrl = "$BaseUrl/Jobs?`$filter=(CreationTime ge $FilterDate)&`$orderby=CreationTime desc&`$top=1000"

    while ($NextUrl) {
        try {
            $Response = Invoke-RestMethod -Uri $NextUrl -Headers $FolderHeaders -Method Get
            if ($Response.value) { $Jobs += $Response.value }
            $NextUrl = $Response.'@odata.nextLink' # Gère la pagination
        } catch {
            # Affichage de l'erreur (y compris 429) et arrêt de la récupération pour ce folder.
            $StatusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { "Inconnu" }
            Write-Host "❌ Erreur (HTTP $StatusCode) pour folder $FolderName. Récupération abandonnée pour ce dossier." -ForegroundColor Red
            $NextUrl = $null 
        }
    }

    foreach ($job in $Jobs) { $job | Add-Member -NotePropertyName FolderName -NotePropertyValue $FolderName -Force }
    Write-Host "📦 [$FolderName] Total jobs récupérés : $($Jobs.Count)"
    return $Jobs
}


# === Fonction pour exporter les jobs vers la feuille de données ===
function Export-JobsToSheet {
    param ([array]$AllJobs, [string]$SheetName)
    try { $ws = $wb.Worksheets.Item($SheetName) } catch { $ws = $wb.Worksheets.Add(); $ws.Name = $SheetName }
    $ws.Cells.Clear()
    $headers = 'Id','ReleaseName','State','StartTime','EndTime','FolderName'
    for ($i=0; $i -lt $headers.Count; $i++) { $ws.Cells.Item(1, $i+1) = $headers[$i] }
    $row=2
    foreach ($job in $AllJobs) {
        $ws.Cells.Item($row,1) = $job.Id
        $ws.Cells.Item($row,2) = $job.ReleaseName
        $ws.Cells.Item($row,3) = $job.State
        $ws.Cells.Item($row,4) = $job.StartTime
        $ws.Cells.Item($row,5) = $job.EndTime
        $ws.Cells.Item($row,6) = $job.FolderName
        $row++
    }
    if ($ws.Columns.Count -ge 7) { $ws.Columns.Item(7).Delete() | Out-Null }
}


# === Fonction pour exporter le résumé (avec calcul du ROI) ===
function Export-Summary {
    param ([array]$AllJobs, [string]$SheetName)

    $AllStates = @("Successful","Faulted","Stopped","Running","Pending","Terminated","Suspended")

    # On compte le total des jobs réussis et TERMINÉS (EndTime non nul) dans TOUS les folders pour la clé de répartition du coût.
    $TotalSuccessfulAllFolders = ($AllJobs | Where-Object { $_.State -eq "Successful" -and $_.EndTime -ne $null }).Count
    if ($TotalSuccessfulAllFolders -eq 0) { $TotalSuccessfulAllFolders = 1 } 

    $FoldersGrouped = $AllJobs | Group-Object FolderName
    $Summary = @()
    foreach ($group in $FoldersGrouped) {
        $Folder = $group.Name
        $Jobs = $group.Group

        $TotalDownloaded = $Jobs.Count
        $CompletedJobs = $Jobs | Where-Object { $_.State -in @("Successful","Faulted","Stopped","Terminated") }
        $TotalCompleted = $CompletedJobs.Count

        # Comptage états
        $StateCounts = @{ }
        foreach ($s in $AllStates) { $StateCounts[$s] = ($Jobs | Where-Object { $_.State -eq $s }).Count }

        $Success = $StateCounts["Successful"]
        $SuccessRate = if ($TotalCompleted -gt 0) { [math]::Round($Success/$TotalCompleted,2) } else { 0 }
        
        # Le calcul du ROI se fait uniquement sur les jobs qui sont Successful ET qui ont un EndTime (donc finis).
        $SuccessfulFinishedJobs = $Jobs | Where-Object { $_.State -eq "Successful" -and $_.EndTime -ne $null }
        $SuccessCountForROI = $SuccessfulFinishedJobs.Count

        # ROI optimisé : coût proportionnel
        $ProportionalCost = $MonthlyRPACost * ($SuccessCountForROI / $TotalSuccessfulAllFolders)
        if ($TotalSuccessfulAllFolders -eq 1) { $ProportionalCost = 0 }

        # Gain humain
        $TotalHoursSaved = [math]::Round(($SuccessCountForROI * $MinutesSavedPerJob)/60,2)
        $HumanEquivalentCost = $TotalHoursSaved * $CostPerHour
        $GainNet = [math]::Round($HumanEquivalentCost - $ProportionalCost,2)
        $ROI = if ($ProportionalCost -ne 0) { [math]::Round($GainNet/$ProportionalCost,2) } else { 0 }

        $Summary += [PSCustomObject]@{
            FolderName = $Folder
            TotalJobs = $TotalDownloaded
            Successful = $StateCounts["Successful"]
            Faulted = $StateCounts["Faulted"]
            Stopped = $StateCounts["Stopped"]
            Running = $StateCounts["Running"]
            Pending = $StateCounts["Pending"]
            Terminated = $StateCounts["Terminated"]
            Suspended = $StateCounts["Suspended"]
            SuccessRate = $SuccessRate
            TotalHoursSaved = $TotalHoursSaved
            GainNet = $GainNet
            ROI = $ROI
        }
    }


# === Export vers Excel ===
    try { $wsSummary = $wb.Worksheets.Item($SheetName) } catch { $wsSummary = $wb.Worksheets.Add(); $wsSummary.Name = $SheetName }
    $wsSummary.Cells.Clear()
    $headersSummary = 'FolderName','TotalJobs','Successful','Faulted','Stopped','Running','Pending','Terminated','Suspended','SuccessRate','TotalHoursSaved','GainNet','ROI'
    for ($i=0; $i -lt $headersSummary.Count; $i++) { $wsSummary.Cells.Item(1,$i+1) = $headersSummary[$i] }
    $row=2
    foreach ($item in $Summary) {
        $col=1
        foreach ($key in $headersSummary) {
            $wsSummary.Cells.Item($row,$col) = $item.$key
            $col++
        }
        $row++
    }
}

# === Périodes ===
$NowUtc = (Get-Date).ToUniversalTime()
$Periods = @{
    # Marge de sécurité de 1 heure pour garantir la couverture complète de la période
    "J1"  = $NowUtc.AddHours(-25)
    "J7"  = $NowUtc.AddHours(-(7*24 + 1))
    "J30" = $NowUtc.AddHours(-(30*24 + 1))
}


# === Excel ===
Write-Host "📊 Initialisation Excel..."
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    if (Test-Path $XlsPath) {
        $wb = $excel.Workbooks.Open($XlsPath)
    } else {
        $wb = $excel.Workbooks.Add()
        $wb.SaveAs($XlsPath,51)
    }
} catch {
    Write-Host "❌ Erreur lors de l'ouverture ou de la création du fichier Excel. Vérifiez que le fichier n'est pas ouvert." -ForegroundColor Red
    exit
}


# === Extraction + export ===
foreach ($period in $Periods.Keys) {
    Write-Host ""
    Write-Host "=== Démarrage de l'extraction pour la période $period ===" -ForegroundColor Cyan
    $AllJobs = @()
    foreach ($folder in $Folders) {
        # CORRECTION 429: Délai de 500ms (0.5s) entre chaque appel par dossier
        Start-Sleep -Milliseconds 500
        $AllJobs += Get-UipathJobsForFolder -FolderId $folder.Id -FolderName $folder.DisplayName -StartDate $Periods[$period]
    }
    Write-Host "Export de $($AllJobs.Count) jobs vers Datas_$period..."
    Export-JobsToSheet -AllJobs $AllJobs -SheetName "Datas_$period"
    Write-Host "Calcul du résumé pour Summary_$period..."
    Export-Summary -AllJobs $AllJobs -SheetName "Summary_$period"
}


# === Sauvegarde et fermeture Excel ===
Write-Host ""
Write-Host "💾 Sauvegarde et fermeture d'Excel..."
$wb.Save()
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null


# === On écrit dans la console que l'export des 6 excel feuilles sont OK ===
Write-Host "✅ Export terminé avec succès ! Feuilles : Datas_J1/J7/J30 et Summary_J1/J7/J30" -ForegroundColor Green