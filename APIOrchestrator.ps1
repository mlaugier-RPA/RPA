<#
# API Orchestrator <=> PowerBI
# Maxime LAUGIER
# APIOrchestrator v3.1
# Export XLSX + Summary avec taux de succès (fraction)
#>

# --- Variables ---
$Org = "extiavqvkelj"
$Tenant = "DefaultTenant"
$XlsPath = "H:\Mon Drive\Reporting PowerBI\UiPathJobs.xlsx"   # Chemin export XLSX
$BaseUrl = "https://cloud.uipath.com/$Org/$Tenant/orchestrator_/odata"

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
$Folders = $FoldersResponse.value | Sort-Object DisplayName -Unique  # Évite les doublons

# --- Filtrage sur 7 derniers jours ---
$NowUtc = (Get-Date).ToUniversalTime()
$FilterDate = $NowUtc.AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ssZ")

# --- Fonction pour récupérer les jobs ---
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
            Write-Host "Erreur de récupération pour $FolderName" -ForegroundColor Red
            break
        }
    }

    foreach ($job in $Jobs) {
        $job | Add-Member -NotePropertyName FolderName -NotePropertyValue $FolderName
    }

    return $Jobs
}

# --- Récupération de tous les jobs ---
$AllJobs = @()
foreach ($folder in $Folders) {
    Write-Host "Récupération des jobs pour le folder : $($folder.DisplayName)"
    $AllJobs += Get-UipathJobsForFolder -FolderId $folder.Id -FolderName $folder.DisplayName
}

# --- Export Excel ---
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

# --- Feuille "Summary" ---
try { $wsSummary = $wb.Worksheets.Item("Summary") } catch { $wsSummary = $wb.Worksheets.Add(); $wsSummary.Name = "Summary" }
$wsSummary.Cells.Clear()

$headersSummary = 'FolderName','TotalJobs','Successful','Faulted','SuccessRate'
for ($i=0; $i -lt $headersSummary.Count; $i++) {
    $wsSummary.Cells.Item(1, $i+1) = $headersSummary[$i]
}

$FoldersSummary = @()

foreach ($folder in $Folders) {
    $JobsInFolder = $AllJobs | Where-Object { $_.FolderName -eq $folder.DisplayName }
    $Total = $JobsInFolder.Count

    # Ignore si aucun job
    if ($Total -eq 0) { continue }

    $Success = ($JobsInFolder | Where-Object {$_.State -eq 'Successful'}).Count
    $Faulted = ($JobsInFolder | Where-Object {$_.State -eq 'Faulted'}).Count

    if ($Total -gt 0) {
        $Taux = [math]::Round(($Success / $Total), 2)
    } else {
        $Taux = $null
    }

    $FoldersSummary += [PSCustomObject]@{
        FolderName  = $folder.DisplayName
        TotalJobs   = $Total
        Successful  = $Success
        Faulted     = $Faulted
        SuccessRate = $Taux
    }
}

# Ajout des totaux si au moins un folder a des jobs
if ($FoldersSummary.Count -gt 0) {
    $TotalJobsTermines = ($AllJobs | Measure-Object).Count
    $SuccessfulJobs = $AllJobs | Where-Object {$_.State -eq 'Successful'}
    $FaultedJobs = $AllJobs | Where-Object {$_.State -eq 'Faulted'}

    $FoldersSummary += [PSCustomObject]@{
        FolderName  = 'TOTAL'
        TotalJobs   = $TotalJobsTermines
        Successful  = $SuccessfulJobs.Count
        Faulted     = $FaultedJobs.Count
        SuccessRate = if ($TotalJobsTermines -gt 0) { [math]::Round(($SuccessfulJobs.Count / $TotalJobsTermines),2) } else { $null }
    }
}

# Écriture dans Excel
$row = 2
foreach ($item in $FoldersSummary) {
    $wsSummary.Cells.Item($row,1) = $item.FolderName
    $wsSummary.Cells.Item($row,2) = $item.TotalJobs
    $wsSummary.Cells.Item($row,3) = $item.Successful
    $wsSummary.Cells.Item($row,4) = $item.Faulted
    $wsSummary.Cells.Item($row,5) = $item.SuccessRate
    $row++
}

# --- Sauvegarde ---
$wb.Save()
$wb.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsSummary) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "✅ Export terminé :"
Write-Host " - Datas : toutes les exécutions des 7 derniers jours"
Write-Host " - Summary : synthèse avec taux de succès (fraction 0-1)"
