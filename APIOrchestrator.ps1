<#
# API Orchestrator <=> PowerBI
# Version : v9.1 
# Maxime LAUGIER
# Update du 13/10/2025
# Période de temps : J-30 jours, J-7days et j-24h
#>

# === variables fixes pour tout le script ===
$Org = "extiavqvkelj"
$Tenant = "DefaultTenant"
$XlsPath = "H:\Mon Drive\Reporting PowerBI\UiPathJobs.xlsx"
$BaseUrl = "https://cloud.uipath.com/$Org/$Tenant/orchestrator_/odata"


# === Paramètres ROI pour le calcul du gain ===
$CostPerHour = 30         # € / heure d’un humain
$MinutesSavedPerJob = 20  # minutes économisées par job réussi
$MonthlyRPACost = 5900    # € coût global RPA mensuel


# === authentification sur l'api et on essaie de récupérer le token ===
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")

$body = "client_id=92615cee-13a8-4195-b52a-3543976033cc&client_secret=lOa%5EtVshMA!mLwLsI8kbwNO)8QH%23p1c%23Qa_jmIN%3FCkYo~YOevEs73EVc(Cb(N2jy&grant_type=client_credentials"

$response = Invoke-RestMethod "https://cloud.uipath.com/$Org/identity_/connect/token" -Method 'POST' -Headers $headers -Body $body
$PAT = $response.access_token
if (-not $PAT) { Write-Host "❌ Token introuvable" -ForegroundColor Red; exit }


# === Après avoir récupérer le token dynamique, on set-up les headers pour le call API ===
$Headers = @{
    "Authorization" = "Bearer $PAT"
    "Accept" = "application/json;odata=nometadata"
}


# === Récupération des folders ===
$Folders = (Invoke-RestMethod -Uri "$BaseUrl/Folders" -Headers $Headers).value
if (-not $Folders) { Write-Host "❌ Aucun folder trouvé"; exit }


# === Fonction pour avoir l'ID des folder dans Robot_Extia d'UIpath Orchestrator et faire un for each pour checker chaque dossier ===
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
    $NextUrl = "$BaseUrl/Jobs?`$filter=(EndTime ge $FilterDate)&`$orderby=EndTime desc&`$top=1000"

    while ($NextUrl) {
        try {
            $Response = Invoke-RestMethod -Uri $NextUrl -Headers $FolderHeaders -Method Get
            if ($Response.value) { $Jobs += $Response.value }
            $NextUrl = $Response.'@odata.nextLink'
        } catch {
            Write-Host "⚠ Erreur pour folder $FolderName" -ForegroundColor Yellow
            break
        }
    }

    foreach ($job in $Jobs) { $job | Add-Member -NotePropertyName FolderName -NotePropertyValue $FolderName }
    Write-Host "📦 [$FolderName] Total jobs récupérés : $($Jobs.Count)"
    return $Jobs
}


# === Fonction pour récupérer les jobs trouver dans la doc API d'UiPath Orchestrator ===
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


# === Fonction pour exporter les jobs dans une sheet spécifique ===
function Export-Summary {
    param ([array]$AllJobs, [string]$SheetName)

    $AllStates = @("Successful","Faulted","Stopped","Running","Pending","Terminated","Suspended")

    $FoldersGrouped = $AllJobs | Group-Object FolderName
    $TotalSuccessfulAllFolders = ($AllJobs | Where-Object { $_.State -eq "Successful" }).Count
    if ($TotalSuccessfulAllFolders -eq 0) { $TotalSuccessfulAllFolders = 1 } # éviter division par zéro

    $Summary = @()
    foreach ($group in $FoldersGrouped) {
        $Folder = $group.Name
        $Jobs = $group.Group

        # Total des jobs terminés (pour le calcul du taux)
        $CompletedJobs = $Jobs | Where-Object { $_.State -in @("Successful","Faulted","Stopped","Terminated") }
        $Total = $CompletedJobs.Count

        # Comptage états
        $StateCounts = @{ }
        foreach ($s in $AllStates) { $StateCounts[$s] = ($Jobs | Where-Object { $_.State -eq $s }).Count }

        $Success = $StateCounts["Successful"]
        $SuccessRate = if ($Total -gt 0) { [math]::Round($Success/$Total,2) } else { 0 }

        # ROI optimisé : coût proportionnel basé uniquement sur les jobs réussis
        $ProportionalCost = $MonthlyRPACost * ($Success / $TotalSuccessfulAllFolders)

        # Gain humain
        $TotalHoursSaved = [math]::Round(($Success * $MinutesSavedPerJob)/60,2)
        $HumanEquivalentCost = $TotalHoursSaved * $CostPerHour
        $GainNet = [math]::Round($HumanEquivalentCost - $ProportionalCost,2)
        $ROI = if ($ProportionalCost -ne 0) { [math]::Round($GainNet/$ProportionalCost,2) } else { 0 }

        $Summary += [PSCustomObject]@{
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
        if ($item.TotalJobs -eq 0) { continue }
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
    "J1"  = $NowUtc.AddDays(-1)
    "J7"  = $NowUtc.AddDays(-7)
    "J30" = $NowUtc.AddDays(-30)
    "J354" = $NowUtc.AddDays(-354)
}


# === Excel ===
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
if (Test-Path $XlsPath) { $wb = $excel.Workbooks.Open($XlsPath) } else { $wb = $excel.Workbooks.Add(); $wb.SaveAs($XlsPath,51) }


# === Extraction + export ===
foreach ($period in $Periods.Keys) {
    $AllJobs = @()
    foreach ($folder in $Folders) {
        Write-Host "📦 [$period] Récupération jobs folder: $($folder.DisplayName)"
        $AllJobs += Get-UipathJobsForFolder -FolderId $folder.Id -FolderName $folder.DisplayName -StartDate $Periods[$period]
    }

    Export-JobsToSheet -AllJobs $AllJobs -SheetName "Datas_$period"
    Export-Summary -AllJobs $AllJobs -SheetName "Summary_$period"
}


# === Sauvegarde et fermeture Excel ===
$wb.Save()
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null


# === On écrit dans la console que l'export des 6 excel feuilles sont OK ===
Write-Host "✅ Export terminé ! Feuilles : Datas_J1/J7/J30 et Summary_J1/J7/J30"
