# =====================================================================
# API Orchestrator <=> PowerBI - Script Finalisé avec LOGS
# Version : v10.0
# FINAL : Ajout des feuilles Logs_J* (Info, Warn, Error, Trace)
# Auteur : Maxime LAUGIER
# Update du 21/10/2025
# =====================================================================


# === variables fixes pour tout le script ===
$Org = "extiavqvkelj"
$Tenant = "DefaultTenant"
$XlsPath = "H:\Mon Drive\Reporting PowerBI\UiPathJobs.xlsx"
$BaseUrl = "https://cloud.uipath.com/$Org/$Tenant/orchestrator_/odata"


# === Paramètres ROI pour le calcul du gain ===
$CostPerHour = 30          # € / heure d’un humain
$MinutesSavedPerJob = 20   # minutes économisées par job réussi
$MonthlyRPACost = 5900     # € coût global RPA mensuel



# === Suppression du fichier Excel existant au début ===
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force

if (Test-Path $XlsPath) {
    try {
        Remove-Item $XlsPath -Force
        Write-Host "🗑️ Ancien fichier Excel supprimé : $XlsPath"
    } catch {
        Write-Host "❌ Impossible de supprimer le fichier Excel : $($_.Exception.Message)" -ForegroundColor Red
        exit
    }
}



# === Authentification ===
Write-Host "🔑 Tentative de récupération du jeton d'accès..."
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")

$body = "client_id=92615cee-13a8-4195-b52a-3543976033cc&client_secret=lOa%5EtVshMA!mLwLsI8kbwNO)8QH%23p1c%23Qa_jmIN%3FCkYo~YOevEs73EVc(Cb(N2jy&grant_type=client_credentials"

try {
    $response = Invoke-RestMethod "https://cloud.uipath.com/$Org/identity_/connect/token" -Method 'POST' -Headers $headers -Body $body
    $PAT = $response.access_token
} catch {
    Write-Host "❌ Erreur d'authentification : $($_.Exception.Message)" -ForegroundColor Red
    exit
}

if (-not $PAT) { Write-Host "❌ Token introuvable" -ForegroundColor Red; exit }
Write-Host "✅ Jeton récupéré."


# === Headers ===
$Headers = @{
    "Authorization" = "Bearer $PAT"
    "Accept" = "application/json;odata=nometadata"
}


# === Récupération des dossiers ===
try {
    $Folders = (Invoke-RestMethod -Uri "$BaseUrl/Folders" -Headers $Headers).value
} catch {
    Write-Host "❌ Erreur lors de la récupération des folders : $($_.Exception.Message)" -ForegroundColor Red
    exit
}
if (-not $Folders) { Write-Host "❌ Aucun folder trouvé" -ForegroundColor Yellow; exit }
Write-Host "📁 $(@($Folders).Count) folders trouvés."


# =====================================================================
# === FONCTIONS ===
# =====================================================================

# === Récupération des jobs ===
function Get-UipathJobsForFolder {
    param (
        [string]$FolderId,
        [string]$FolderName,
        [datetime]$StartDate
    )

    $FolderHeaders = @{}
    foreach ($key in $Headers.Keys) { $FolderHeaders[$key] = $Headers[$key] }
    $FolderHeaders["X-UIPATH-OrganizationUnitId"] = "$FolderId"

    $Jobs = @()
    $FilterDate = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $NextUrl = "$BaseUrl/Jobs?`$filter=(CreationTime ge $FilterDate)&`$orderby=CreationTime desc&`$top=1000"

    while ($NextUrl) {
        try {
            $Response = Invoke-RestMethod -Uri $NextUrl -Headers $FolderHeaders -Method Get
            if ($Response.value) { $Jobs += $Response.value }
            $NextUrl = $Response.'@odata.nextLink'
            if ($NextUrl) { Start-Sleep -Milliseconds 200 }
        } catch {
            $StatusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { "Inconnu" }
            Write-Host "❌ Erreur (HTTP $StatusCode) pour folder $FolderName." -ForegroundColor Red
            $NextUrl = $null
        }
    }

    foreach ($job in $Jobs) { $job | Add-Member -NotePropertyName FolderName -NotePropertyValue $FolderName -Force }
    Write-Host "📦 [$FolderName] Total jobs récupérés : $($Jobs.Count)"
    return $Jobs
}



# === Récupération des logs ===
function Get-UipathLogsForFolder {
    param (
        [string]$FolderId,
        [string]$FolderName,
        [datetime]$StartDate
    )

    $FolderHeaders = @{}
    foreach ($key in $Headers.Keys) { $FolderHeaders[$key] = $Headers[$key] }
    $FolderHeaders["X-UIPATH-OrganizationUnitId"] = "$FolderId"

    $Logs = @()
    $FilterDate = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $NextUrl = "$BaseUrl/RobotLogs?`$filter=(TimeStamp ge $FilterDate)&`$orderby=TimeStamp desc&`$top=1000"

    while ($NextUrl) {
        try {
            $Response = Invoke-RestMethod -Uri $NextUrl -Headers $FolderHeaders -Method Get
            if ($Response.value) { $Logs += $Response.value }
            $NextUrl = $Response.'@odata.nextLink'
            if ($NextUrl) { Start-Sleep -Milliseconds 200 }
        } catch {
            $StatusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { "Inconnu" }
            Write-Host "❌ Erreur (HTTP $StatusCode) pour les logs du dossier $FolderName." -ForegroundColor Red
            $NextUrl = $null
        }
    }

    foreach ($log in $Logs) { $log | Add-Member -NotePropertyName FolderName -NotePropertyValue $FolderName -Force }
    Write-Host "📝 [$FolderName] Logs récupérés : $($Logs.Count)"
    return $Logs
}



# === Export Jobs avec ROI ===
function Export-JobsToSheet {
    param (
        [array]$AllJobs, 
        [string]$SheetName,
        [array]$SummaryData
    )
    
    $SummaryLookup = @{}
    $SummaryData | ForEach-Object { $SummaryLookup[$_.FolderName] = $_ }
    
    try { $ws = $wb.Worksheets.Item($SheetName) } catch { $ws = $wb.Worksheets.Add(); $ws.Name = $SheetName }
    $ws.Cells.Clear()

    $headers = 'Id','ReleaseName','State','StartTime','EndTime','FolderName'
    for ($i=0; $i -lt $headers.Count; $i++) { $ws.Cells.Item(1, $i+1) = $headers[$i] }

    $headersROI = 'SuccessRate','TotalHoursSaved','GainNet','ROI'
    $col = $headers.Count + 1
    foreach ($h in $headersROI) {
        $ws.Cells.Item(1, $col) = $h
        $col++
    }

    $ws.Range("A1:J1").Font.Bold = $true
    $row=2
    foreach ($job in $AllJobs) {
        $ws.Cells.Item($row,1) = $job.Id
        $ws.Cells.Item($row,2) = $job.ReleaseName
        $ws.Cells.Item($row,3) = $job.State
        $ws.Cells.Item($row,4) = $job.StartTime
        $ws.Cells.Item($row,5) = $job.EndTime
        $ws.Cells.Item($row,6) = $job.FolderName
        $summaryItem = $SummaryLookup[$job.FolderName]
        if ($summaryItem) {
            $ws.Cells.Item($row,7) = $summaryItem.SuccessRate
            $ws.Cells.Item($row,8) = $summaryItem.TotalHoursSaved
            $ws.Cells.Item($row,9) = $summaryItem.GainNet
            $ws.Cells.Item($row,10) = $summaryItem.ROI
        }
        $row++
    }
    $ws.Columns.AutoFit() | Out-Null
}


# === Export Summary ===
function Export-SummaryToSheet {
    param ([array]$SummaryData, [string]$SheetName)
    
    try { $ws = $wb.Worksheets.Item($SheetName) } catch { $ws = $wb.Worksheets.Add(); $ws.Name = $SheetName }
    $ws.Cells.Clear()
    
    $headersSummary = 'FolderName','TotalJobs','Successful','Faulted','Stopped','Running','Pending','Terminated','Suspended','Waiting','Stopping','SuccessRate','TotalHoursSaved','GainNet','ROI'
    for ($i=0; $i -lt $headersSummary.Count; $i++) { $ws.Cells.Item(1,$i+1) = $headersSummary[$i] }
    $row=2
    foreach ($item in $SummaryData) {
        $col=1
        foreach ($key in $headersSummary) {
            $ws.Cells.Item($row,$col) = $item.$key
            $col++
        }
        $row++
    }
    $ws.Columns.AutoFit() | Out-Null
}


# === Export Logs ===
function Export-LogsToSheet {
    param (
        [array]$AllLogs,
        [string]$SheetName
    )

    try { $ws = $wb.Worksheets.Item($SheetName) } catch { $ws = $wb.Worksheets.Add(); $ws.Name = $SheetName }
    $ws.Cells.Clear()

    $headers = 'TimeStamp','Level','Message','JobKey','ProcessName','MachineName','FolderName'
    for ($i=0; $i -lt $headers.Count; $i++) { $ws.Cells.Item(1, $i+1) = $headers[$i] }

    $ws.Range("A1:G1").Font.Bold = $true

    $row = 2
    foreach ($log in $AllLogs) {
        $ws.Cells.Item($row,1) = $log.TimeStamp
        $ws.Cells.Item($row,2) = $log.Level
        $ws.Cells.Item($row,3) = $log.Message
        $ws.Cells.Item($row,4) = $log.JobKey
        $ws.Cells.Item($row,5) = $log.ProcessName
        $ws.Cells.Item($row,6) = $log.MachineName
        $ws.Cells.Item($row,7) = $log.FolderName
        $row++
    }

    $ws.Columns.AutoFit() | Out-Null

    # Mise en forme des erreurs (rouge) et warnings (orange)
    $UsedRange = $ws.UsedRange
    $UsedRange.AutoFilter()
    $ws.Range("A1").AutoFilter(2, "Error")
    $ws.Range("A1").AutoFilter(2, "Warn")
}


# === Calcul du résumé ROI ===
function Export-Summary {
    param ([array]$AllJobs)

    $AllStates = @("Successful","Faulted","Stopped","Running","Pending","Terminated","Suspended","Waiting","Stopping")
    $TotalSuccessfulAllFolders = ($AllJobs | Where-Object { $_.State -eq "Successful" -and $_.EndTime -ne $null }).Count
    if ($TotalSuccessfulAllFolders -eq 0) { $TotalSuccessfulAllFolders = 1 }

    $Summary = @()
    $FoldersGrouped = $AllJobs | Group-Object FolderName
    foreach ($group in $FoldersGrouped) {
        $Folder = $group.Name
        $Jobs = $group.Group
        if ($Jobs.Count -eq 0) { continue }

        $StateCounts = @{}
        foreach ($s in $AllStates) { $StateCounts[$s] = ($Jobs | Where-Object { $_.State -eq $s }).Count }

        $Success = $StateCounts["Successful"]
        $Completed = ($Jobs | Where-Object { $_.State -in @("Successful","Faulted","Stopped","Terminated") }).Count
        $SuccessRate = if ($Completed -gt 0) { [math]::Round($Success/$Completed,2) } else { 0 }

        $SuccessfulFinishedJobs = $Jobs | Where-Object { $_.State -eq "Successful" -and $_.EndTime -ne $null }
        $SuccessCountForROI = $SuccessfulFinishedJobs.Count
        $ProportionalCost = $MonthlyRPACost * ($SuccessCountForROI / $TotalSuccessfulAllFolders)
        $TotalHoursSaved = [math]::Round(($SuccessCountForROI * $MinutesSavedPerJob)/60,2)
        $HumanEquivalentCost = $TotalHoursSaved * $CostPerHour
        $GainNet = [math]::Round($HumanEquivalentCost - $ProportionalCost,2)
        $ROI = if ($ProportionalCost -ne 0) { [math]::Round($GainNet/$ProportionalCost,2) } else { 0 }

        $Summary += [PSCustomObject]@{
            FolderName = $Folder
            TotalJobs = $Jobs.Count
            Successful = $StateCounts["Successful"]
            Faulted = $StateCounts["Faulted"]
            Stopped = $StateCounts["Stopped"]
            Running = $StateCounts["Running"]
            Pending = $StateCounts["Pending"]
            Terminated = $StateCounts["Terminated"]
            Suspended = $StateCounts["Suspended"]
            Waiting = $StateCounts["Waiting"]
            Stopping = $StateCounts["Stopping"]
            SuccessRate = $SuccessRate
            TotalHoursSaved = $TotalHoursSaved
            GainNet = $GainNet
            ROI = $ROI
        }
    }
    return $Summary
}


# =====================================================================
# === PÉRIODES & EXCEL ===
# =====================================================================

$NowUtc = (Get-Date).ToUniversalTime()
$Periods = @{
    "J1"  = $NowUtc.AddHours(-25)
    "J7"  = $NowUtc.AddHours(-(7*24 + 1))
    "J30" = $NowUtc.AddHours(-(30*24 + 1))
}


Write-Host "📊 Initialisation Excel..."
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($XlsPath,51)
} catch {
    Write-Host "❌ Erreur ouverture Excel." -ForegroundColor Red
    exit
}


# =====================================================================
# === BOUCLE PRINCIPALE ===
# =====================================================================
foreach ($period in $Periods.Keys) {
    Write-Host ""
    Write-Host "=== Extraction pour $period ===" -ForegroundColor Cyan
    $AllJobs = @()
    $AllLogs = @()

    foreach ($folder in $Folders) {
        Start-Sleep -Milliseconds 500
        $AllJobs += Get-UipathJobsForFolder -FolderId $folder.Id -FolderName $folder.DisplayName -StartDate $Periods[$period]
        $AllLogs += Get-UipathLogsForFolder -FolderId $folder.Id -FolderName $folder.DisplayName -StartDate $Periods[$period]
    }

    $SummaryData = Export-Summary -AllJobs $AllJobs

    Write-Host "Export Datas_$period..."
    Export-JobsToSheet -AllJobs $AllJobs -SheetName "Datas_$period" -SummaryData $SummaryData

    Write-Host "Export Summary_$period..."
    Export-SummaryToSheet -SummaryData $SummaryData -SheetName "Summary_$period"

    Write-Host "Export Logs_$period..."
    Export-LogsToSheet -AllLogs $AllLogs -SheetName "Logs_$period"
}


# =====================================================================
# === FIN ===
# =====================================================================
Write-Host ""
Write-Host "💾 Sauvegarde du fichier Excel..."
$wb.Save()
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "✅ Export terminé avec succès : Datas_J*/Summary_J*/Logs_J*" -ForegroundColor Green
