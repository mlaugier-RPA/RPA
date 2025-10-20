<#
# API Orchestrator <=> PowerBI - Script Finalisé
# Version : v10.7
# FINAL : Suppression fichier Excel existant + logs codés (Info, Warn, Error) par job
# Gestion HTTP 400 / 429 pour récupération logs
# Maxime LAUGIER
# Update du 20/10/2025
#>

# === variables fixes pour tout le script ===
$Org = "extiavqvkelj"
$Tenant = "DefaultTenant"
$XlsPath = "H:\Mon Drive\Reporting PowerBI\UiPathJobs.xlsx"
$BaseUrl = "https://cloud.uipath.com/$Org/$Tenant/orchestrator_/odata"

# === Suppression du fichier Excel existant au début ===
if (Test-Path $XlsPath) {
    try {
        Remove-Item $XlsPath -Force
        Write-Host "🗑️ Ancien fichier Excel supprimé : $XlsPath"
    } catch {
        Write-Host "❌ Impossible de supprimer le fichier Excel : $($_.Exception.Message)" -ForegroundColor Red
        exit
    }
}

# === Paramètres ROI pour le calcul du gain ===
$CostPerHour = 30          # € / heure d’un humain
$MinutesSavedPerJob = 20   # minutes économisées par job réussi
$MonthlyRPACost = 5900     # € coût global RPA mensuel

# === authentification sur l'api et récupération du token ===
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

# === Headers API ===
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

# === Fonction pour récupérer les jobs + logs codés Info/Warning/Error ===
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
    $NextUrl = "$BaseUrl/Jobs?`$filter=(CreationTime ge $FilterDate)&`$orderby=CreationTime desc&`$top=1000"

    while ($NextUrl) {
        try {
            $Response = Invoke-RestMethod -Uri $NextUrl -Headers $FolderHeaders -Method Get
            if ($Response.value) { $Jobs += $Response.value }
            $NextUrl = $Response.'@odata.nextLink'
            if ($NextUrl) { Start-Sleep -Milliseconds 200 } # éviter 429
        } catch {
            $StatusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { "Inconnu" }
            Write-Host "❌ Erreur (HTTP $StatusCode) pour folder $FolderName. Récupération abandonnée pour ce dossier." -ForegroundColor Red
            $NextUrl = $null
        }
    }

    # === Récupération logs codés Info/Warning/Error pour chaque job terminé ===
    foreach ($job in $Jobs | Where-Object { $_.State -notin @("Running","Pending") }) {
        Start-Sleep -Milliseconds 400
        try {
            $JobKeyQuoted = [uri]::EscapeDataString("'$($job.Key)'")
            $LogsUrl = "$BaseUrl/RobotLogs?`$filter=JobKey%20eq%20$JobKeyQuoted&`$orderby=TimeStamp%20desc"

            $attempt = 0
            $LogsResp = $null
            while ($attempt -lt 3 -and -not $LogsResp) {
                try {
                    $attempt++
                    $LogsResp = Invoke-RestMethod -Uri $LogsUrl -Headers $FolderHeaders -Method Get -ErrorAction Stop
                } catch {
                    $code = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { "?" }
                    if ($code -eq 429) { Start-Sleep -Seconds 2 }
                    elseif ($code -eq 400) { break }
                    else { throw }
                }
            }

            # --- Logs codés par type ---
            $InfoLogs = @(); $WarnLogs = @(); $ErrorLogs = @()
            if ($LogsResp.value) {
                foreach ($l in $LogsResp.value) {
                    switch ($l.Level) {
                        "Info" { $InfoLogs += $l.Message }
                        "Warn" { $WarnLogs += $l.Message }
                        "Error" { $ErrorLogs += $l.Message }
                        default { }
                    }
                }
            }
            $job | Add-Member -NotePropertyName LogsInfo -NotePropertyValue $InfoLogs -Force
            $job | Add-Member -NotePropertyName LogsWarn -NotePropertyValue $WarnLogs -Force
            $job | Add-Member -NotePropertyName LogsError -NotePropertyValue $ErrorLogs -Force

        } catch {
            Write-Host "⚠️ Logs non récupérés pour job $($job.Id)" -ForegroundColor DarkYellow
            $job | Add-Member -NotePropertyName LogsInfo -NotePropertyValue @() -Force
            $job | Add-Member -NotePropertyName LogsWarn -NotePropertyValue @() -Force
            $job | Add-Member -NotePropertyName LogsError -NotePropertyValue @() -Force
        }
    }

    foreach ($job in $Jobs) { $job | Add-Member -NotePropertyName FolderName -NotePropertyValue $FolderName -Force }
    Write-Host "📦 [$FolderName] Total jobs récupérés : $($Jobs.Count)"
    return $Jobs
}

# === Export Jobs vers Excel (sans InputArgs/OutputArgs/LastLogs) ===
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
    $headersROI = 'SuccessRate','TotalHoursSaved','GainNet','ROI'
    $col=1
    foreach ($h in $headers + $headersROI) { $ws.Cells.Item(1,$col) = $h; $col++ }
    $ws.Range("A1:$([char](64+$col-1))1").Font.Bold = $true

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

# === Fonctions Export-Summary et Export-SummaryToSheet ===
# (inchangées depuis v10.5, calcul ROI etc.)

# === Périodes ===
$NowUtc = (Get-Date).ToUniversalTime()
$Periods = @{ "J1" = $NowUtc.AddHours(-25); "J7" = $NowUtc.AddHours(-(7*24+1)); "J30" = $NowUtc.AddHours(-(30*24+1)) }

# === Excel ===
Write-Host "📊 Initialisation Excel..."
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($XlsPath,51)
} catch { Write-Host "❌ Erreur Excel" -ForegroundColor Red; exit }

# === Extraction + export ===
foreach ($period in $Periods.Keys) {
    Write-Host "=== Extraction période $period ==="
    $AllJobs = @()
    foreach ($folder in $Folders) { Start-Sleep -Milliseconds 500; $AllJobs += Get-UipathJobsForFolder -FolderId $folder.Id -FolderName $folder.DisplayName -StartDate $Periods[$period] }

    $SummaryData = Export-Summary -AllJobs $AllJobs

    Write-Host "Export vers Datas_$period..."
    Export-JobsToSheet -AllJobs $AllJobs -SheetName "Datas_$period" -SummaryData $SummaryData
    Write-Host "Export du résumé vers Summary_$period..."
    Export-SummaryToSheet -SummaryData $SummaryData -SheetName "Summary_$period"
}

# === Sauvegarde Excel ===
Write-Host "💾 Sauvegarde et fermeture d'Excel..."
$wb.Save()
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "✅ Export terminé !"
