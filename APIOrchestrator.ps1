<#
# API Orchestrator <=> PowerBI
# Maxime LAUGIER
# APIOrchestrator v3.7
# Export XLSX + Summary + ROI relatif + GainNet + 30 jours + résumé Datas + colonne G supprimée
# Gestion des valeurs vides pour TotalJobs, Successful, Faulted
#>

# --- Variables générales ---
$Org = "extiavqvkelj"
$Tenant = "DefaultTenant"
$XlsPath = "H:\Mon Drive\Reporting PowerBI\UiPathJobs.xlsx"
$BaseUrl = "https://cloud.uipath.com/$Org/$Tenant/orchestrator_/odata"

# --- Paramètres ROI ---
$CostPerHour = 30         # € / heure d’un humain
$MinutesSavedPerJob = 15  # Temps gagné par job réussi (en minutes)
$MonthlyRPACost = 5900    # Coût global RPA mensuel (€)

# --- Authentification ---
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")

$body = "client_id=92615cee-13a8-4195-b52a-3543976033cc&client_secret=lOa%5EtVshMA!mLwLsI8kbwNO)8QH%23p1c%23Qa_jmIN%3FCkYo~YOevEs73EVc(Cb(N2jy&grant_type=client_credentials"

$response = Invoke-RestMethod "https://cloud.uipath.com/$Org/identity_/connect/token" -Method 'POST' -Headers $headers -Body $body
$PAT = $response.access_token

$Headers = @{
    "Authorization" = "Bearer $PAT"
    "Accept" = "application/json;odata=nometadata"
}

# --- Récupération des folders ---
$FoldersResponse = Invoke-RestMethod -Uri "$BaseUrl/Folders" -Headers $Headers -Method Get
$Folders = $FoldersResponse.value | Sort-Object DisplayName -Unique

# --- Filtrage sur 30 derniers jours ---
$NowUtc = (Get-Date).ToUniversalTime()
$FilterDate = $NowUtc.AddDays(-30).ToString("yyyy-MM-ddTHH:mm:ssZ")

# --- Fonction : récupération des jobs par folder ---
function Get-UipathJobsForFolder {
    param (
        [Parameter(Mandatory=$true)] [string]$FolderId,
        [Parameter(Mandatory=$true)] [string]$FolderName
    )

    $FolderHeaders = $Headers.Clone()
    $FolderHeaders.Add("X-UIPATH-OrganizationUnitId", $FolderId)

    $Jobs = @()
    $NextUrl = "$BaseUrl/Jobs?`$filter=(EndTime ge $FilterDate) and (State eq 'Successful' or State eq 'Faulted')&`$orderby=EndTime desc&`$top=100"

    while ($NextUrl) {
        try {
            $Response = Invoke-RestMethod -Uri $NextUrl -Headers $FolderHeaders -Method Get
            $Jobs += $Response.value
            $NextUrl = $Response.'@odata.nextLink'
        } catch {
            Write-Host "⚠ Erreur de récupération pour le folder $FolderName" -ForegroundColor Red
            break
        }
    }

    foreach ($job in $Jobs) {
        $job | Add-Member -NotePropertyName FolderName -NotePropertyValue $FolderName
    }

    return $Jobs
}

# --- Récupération globale des jobs ---
$AllJobs = @()
foreach ($folder in $Folders) {
    Write-Host "📥 Récupération des jobs pour : $($folder.DisplayName)"
    $AllJobs += Get-UipathJobsForFolder -FolderId $folder.Id -FolderName $folder.DisplayName
}

# --- Préparation Excel ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

if (Test-Path $XlsPath) {
    $wb = $excel.Workbooks.Open($XlsPath)
} else {
    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($XlsPath, 51)
}

# --- Feuille "Datas" ---
try { $ws = $wb.Worksheets.Item("Datas") } catch { $ws = $wb.Worksheets.Add(); $ws.Name = "Datas" }
$ws.Cells.Clear()

$headers = 'Id','ReleaseName','State','StartTime','EndTime','FolderName'
for ($i=0; $i -lt $headers.Count; $i++) {
    $ws.Cells.Item(1, $i+1) = $headers[$i]
}

$row = 2
foreach ($job in $AllJobs) {
    $ws.Cells.Item($row,1) = $job.Id
    $ws.Cells.Item($row,2) = $job.ReleaseName
    $ws.Cells.Item($row,3) = $job.State
    $ws.Cells.Item($row,4) = $job.StartTime
    $ws.Cells.Item($row,5) = $job.EndTime
    $ws.Cells.Item($row,6) = $job.FolderName
    $row++
}

# --- Supprimer la colonne G inutile ---
$ws.Columns.Item(7).Delete()

# --- Résumé global succès / échecs dans Datas ---
$SuccessCount = ($AllJobs | Where-Object {$_.State -eq 'Successful'}).Count
$FaultedCount = ($AllJobs | Where-Object {$_.State -eq 'Faulted'}).Count
$TotalCount = $SuccessCount + $FaultedCount

$ws.Cells.Item(1,8) = "TotalJobs"
$ws.Cells.Item(1,9) = "Successful"
$ws.Cells.Item(1,10) = "Faulted"

$ws.Cells.Item(2,8) = $TotalCount
$ws.Cells.Item(2,9) = $SuccessCount
$ws.Cells.Item(2,10) = $FaultedCount

# --- Feuille "Summary" ---
try { $wsSummary = $wb.Worksheets.Item("Summary") } catch { $wsSummary = $wb.Worksheets.Add(); $wsSummary.Name = "Summary" }
$wsSummary.Cells.Clear()

$headersSummary = 'FolderName','TotalJobs','Successful','Faulted','SuccessRate','TotalHoursSaved','ROI','GainNet'
for ($i=0; $i -lt $headersSummary.Count; $i++) {
    $wsSummary.Cells.Item(1, $i+1) = $headersSummary[$i]
}

$FoldersSummary = @()

foreach ($folder in $Folders) {
    $JobsInFolder = $AllJobs | Where-Object {
        ($_.FolderName -replace '\s','') -ieq ($folder.DisplayName -replace '\s','')
    }

    if ($JobsInFolder.Count -eq 0) { continue }   # ignore folders sans jobs

    # --- Comptage robuste même si certaines valeurs sont vides ---
    $Success = ($JobsInFolder | Where-Object {$_.State -eq 'Successful'}).Count
    $Faulted = ($JobsInFolder | Where-Object {$_.State -eq 'Faulted'}).Count

    if (-not $Success) { $Success = 0 }
    if (-not $Faulted) { $Faulted = 0 }

    $Total = $Success + $Faulted

    $Taux = if ($Total -gt 0) { [math]::Round(($Success / $Total), 2) } else { $null }

    $TotalMinutesSaved = $Success * $MinutesSavedPerJob
    $TotalHoursSaved = [math]::Round(($TotalMinutesSaved / 60), 2)
    $TotalValue = $TotalHoursSaved * $CostPerHour

    # ROI relatif (1 = seuil rentabilité)
    $ROI = if ($MonthlyRPACost -gt 0) { [math]::Round(($TotalValue / $MonthlyRPACost), 2) } else { $null }
    $GainNet = [math]::Round(($TotalValue - $MonthlyRPACost),2)

    $FoldersSummary += [PSCustomObject]@{
        FolderName      = $folder.DisplayName
        TotalJobs       = $Total
        Successful      = $Success
        Faulted         = $Faulted
        SuccessRate     = $Taux
        TotalHoursSaved = $TotalHoursSaved
        ROI             = $ROI
        GainNet         = $GainNet
    }
}

# --- Totaux globaux ---
if ($FoldersSummary.Count -gt 0) {
    $TotalJobsTermines = ($AllJobs | Measure-Object).Count
    $SuccessfulJobs = $AllJobs | Where-Object {$_.State -eq 'Successful'}
    $FaultedJobs = $AllJobs | Where-Object {$_.State -eq 'Faulted'}

    $TotalMinutesSaved = $SuccessfulJobs.Count * $MinutesSavedPerJob
    $TotalHoursSaved = [math]::Round(($TotalMinutesSaved / 60), 2)
    $TotalValue = $TotalHoursSaved * $CostPerHour
    $ROIglobal = if ($MonthlyRPACost -gt 0) { [math]::Round(($TotalValue / $MonthlyRPACost), 2) } else { $null }
    $GainNetGlobal = [math]::Round(($TotalValue - $MonthlyRPACost),2)

    $FoldersSummary += [PSCustomObject]@{
        FolderName      = 'TOTAL'
        TotalJobs       = $TotalJobsTermines
        Successful      = $SuccessfulJobs.Count
        Faulted         = $FaultedJobs.Count
        SuccessRate     = if ($TotalJobsTermines -gt 0) { [math]::Round(($SuccessfulJobs.Count / $TotalJobsTermines),2) } else { $null }
        TotalHoursSaved = $TotalHoursSaved
        ROI             = $ROIglobal
        GainNet         = $GainNetGlobal
    }
}

# --- Écriture dans Summary ---
$row = 2
foreach ($item in $FoldersSummary) {
    $wsSummary.Cells.Item($row,1) = $item.FolderName
    $wsSummary.Cells.Item($row,2) = $item.TotalJobs
    $wsSummary.Cells.Item($row,3) = $item.Successful
    $wsSummary.Cells.Item($row,4) = $item.Faulted
    $wsSummary.Cells.Item($row,5).NumberFormat = "0.00"
    $wsSummary.Cells.Item($row,5).Value2 = [double]$item.SuccessRate
    $wsSummary.Cells.Item($row,6).NumberFormat = "0.00"
    $wsSummary.Cells.Item($row,6).Value2 = [double]$item.TotalHoursSaved
    $wsSummary.Cells.Item($row,7).NumberFormat = "0.00"
    $wsSummary.Cells.Item($row,7).Value2 = [double]$item.ROI
    $wsSummary.Cells.Item($row,8).NumberFormat = "0.00"
    $wsSummary.Cells.Item($row,8).Value2 = [double]$item.GainNet
    $row++
}

# --- Sauvegarde et nettoyage ---
$wb.Save()
$wb.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsSummary) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "✅ Export terminé :"
Write-Host " - Feuille 'Datas' : toutes les exécutions des 30 derniers jours + résumé succès/échecs, colonne G supprimée"
Write-Host " - Feuille 'Summary' : synthèse nettoyée, uniquement folders avec jobs + taux de succès + ROI + GainNet"
