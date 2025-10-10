<#
# API Orchestrator <=> PowerBI
# Maxime LAUGIER
# Version : v6.3
# Export complet : J-1, J-7, J-30 (Datas + Summary)
#>

# === CONFIG ===
$Org = "extiavqvkelj"
$Tenant = "DefaultTenant"
$XlsPath = "H:\Mon Drive\Reporting PowerBI\UiPathJobs.xlsx"
$BaseUrl = "https://cloud.uipath.com/$Org/$Tenant/orchestrator_/odata"

# --- Paramètres ROI ---
$CostPerHour = 30           # € / heure d’un humain
$MinutesSavedPerJob = 15    # minutes économisées par job réussi
$MonthlyRPACost = 5900      # € coût global RPA mensuel

# === AUTHENTIFICATION ===
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")

$body = "client_id=92615cee-13a8-4195-b52a-3543976033cc&client_secret=lOa%5EtVshMA!mLwLsI8kbwNO)8QH%23p1c%23Qa_jmIN%3FCkYo~YOevEs73EVc(Cb(N2jy&grant_type=client_credentials"

$response = Invoke-RestMethod "https://cloud.uipath.com/$Org/identity_/connect/token" -Method 'POST' -Headers $headers -Body $body
$PAT = $response.access_token

if (-not $PAT) {
    Write-Host "❌ Erreur : Token UiPath introuvable. Vérifie client_id / secret." -ForegroundColor Red
    exit
}

$Headers = @{
    "Authorization" = "Bearer $PAT"
    "Accept" = "application/json;odata=nometadata"
}

# === Fonction principale : récupération des jobs ===
function Get-UipathJobsForFolder {
    param (
        [string]$FolderId,
        [string]$FolderName,
        [datetime]$StartDate
    )

    # Headers spécifiques folder
    $FolderHeaders = @{}
    foreach ($key in $Headers.Keys) { $FolderHeaders[$key] = $Headers[$key] }
    $FolderHeaders["X-UIPATH-OrganizationUnitId"] = "$FolderId"

    $Jobs = @()
    $FilterDate = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $NextUrl = "$BaseUrl/Jobs?`$filter=(EndTime ge $FilterDate)&`$orderby=EndTime desc&`$top=100"

    while ($NextUrl) {
        try {
            $Response = Invoke-RestMethod -Uri $NextUrl -Headers $FolderHeaders -Method Get
            if ($Response.value) { $Jobs += $Response.value }
            $NextUrl = $Response.'@odata.nextLink'
        } catch {
            Write-Host "⚠ Erreur pour le folder $FolderName" -ForegroundColor Yellow
            break
        }
    }

    foreach ($job in $Jobs) {
        $job | Add-Member -NotePropertyName FolderName -NotePropertyValue $FolderName
    }

    return $Jobs
}

# === Récupération des folders ===
$FoldersResponse = Invoke-RestMethod -Uri "$BaseUrl/Folders" -Headers $Headers -Method Get
$Folders = $FoldersResponse.value

# === Fonction : export d'une feuille Excel de données brutes ===
function Export-JobsToSheet {
    param (
        [array]$AllJobs,
        [string]$SheetName
    )

    try {
        $ws = $wb.Worksheets.Item($SheetName)
    } catch {
        $ws = $wb.Worksheets.Add()
        $ws.Name = $SheetName
    }

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

    # Suppression colonne G si existe
    if ($ws.Columns.Count -ge 7) {
        $ws.Columns.Item(7).Delete() | Out-Null
    }
}

# === Fonction : export du résumé agrégé ===
function Export-Summary {
    param (
        [array]$AllJobs,
        [string]$SheetName
    )

    # Tous les états connus
    $AllStates = @("Successful", "Faulted", "Stopped", "Running", "Pending", "Terminated", "Suspended")

    # Création du résumé par Folder
    $Summary = $AllJobs | Group-Object FolderName | ForEach-Object {
        $Folder = $_.Name
        $Jobs = $_.Group
        $Total = $Jobs.Count

        # Comptage des états (même à 0)
        $StateCounts = @{}
        foreach ($s in $AllStates) {
            $StateCounts[$s] = ($Jobs | Where-Object { $_.State -eq $s }).Count
        }

        $Success = $StateCounts["Successful"]
        $Faulted = $StateCounts["Faulted"]
        $SuccessRate = if ($Total -gt 0) { [math]::Round($Success / $Total, 2) } else { 0 }
        if ($SuccessRate -gt 1) { $SuccessRate = 1 }

        $TotalHoursSaved = [math]::Round(($Success * $MinutesSavedPerJob) / 60, 2)
        $HumanEquivalentCost = $TotalHoursSaved * $CostPerHour
        $ROI = [math]::Round(($HumanEquivalentCost - $MonthlyRPACost) / $MonthlyRPACost, 2)
        $GainNet = [math]::Round(($HumanEquivalentCost - $MonthlyRPACost), 2)

        [PSCustomObject]@{
            FolderName = $Folder
            TotalJobs = $Total
            Successful = $StateCounts["Successful"]
            Faulted = $StateCounts["Faulted"]
            Stopped = $StateCounts["Stopped"]
            Running = $StateCounts["Running"]
            Pending = $StateCounts["Pending"]
            Terminated = $StateCounts["Terminated"]
            Suspended = $StateCounts["Suspended"]
            SuccessRate = $SuccessRate
            TotalHoursSaved = $TotalHoursSaved
            ROI = $ROI
            GainNet = $GainNet
        }
    }

    try {
        $wsSummary = $wb.Worksheets.Item($SheetName)
    } catch {
        $wsSummary = $wb.Worksheets.Add()
        $wsSummary.Name = $SheetName
    }

    $wsSummary.Cells.Clear()
    $headersSummary = 'FolderName','TotalJobs','Successful','Faulted','Stopped','Running','Pending','Terminated','Suspended','SuccessRate','TotalHoursSaved','ROI','GainNet'
    for ($i=0; $i -lt $headersSummary.Count; $i++) {
        $wsSummary.Cells.Item(1, $i+1) = $headersSummary[$i]
    }

    $row = 2
    foreach ($item in $Summary) {
        if ($item.TotalJobs -eq 0) { continue }
        $col = 1
        foreach ($key in $headersSummary) {
            $wsSummary.Cells.Item($row, $col) = $item.$key
            $col++
        }
        $row++
    }
}

# === Périodes à extraire ===
$NowUtc = (Get-Date).ToUniversalTime()
$Periods = @{
    "J1"  = $NowUtc.AddDays(-1)
    "J7"  = $NowUtc.AddDays(-7)
    "J30" = $NowUtc.AddDays(-30)
}

# === Initialisation Excel ===
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

if (Test-Path $XlsPath) {
    $wb = $excel.Workbooks.Open($XlsPath)
} else {
    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($XlsPath, 51)
}

# === Récupération et export pour chaque période ===
foreach ($period in $Periods.Keys) {
    $AllJobs = @()
    foreach ($folder in $Folders) {
        Write-Host "📦 [$period] Récupération jobs folder: $($folder.DisplayName)"
        $AllJobs += Get-UipathJobsForFolder -FolderId $folder.Id -FolderName $folder.DisplayName -StartDate $Periods[$period]
    }

    Export-JobsToSheet -AllJobs $AllJobs -SheetName "Datas_$period"
    Export-Summary -AllJobs $AllJobs -SheetName "Summary_$period"
}

# === Sauvegarde et nettoyage ===
$wb.Save()
$wb.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "✅ Export terminé !"
Write-Host "Feuilles générées : Datas_J1 / J7 / J30 + Summary_J1 / J7 / J30"
