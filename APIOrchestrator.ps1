<#
# API Orchestrator <=> PowerBI
# Maxime LAUGIER
# APIOrchestrator v5.0
# Multi-feuilles + Comparaison J-30 vs J-7
#>

# --- Variables générales ---
$Org = "extiavqvkelj"
$Tenant = "DefaultTenant"
$XlsPath = "H:\Mon Drive\Reporting PowerBI\UiPathJobs.xlsx"
$BaseUrl = "https://cloud.uipath.com/$Org/$Tenant/orchestrator_/odata"

# --- Paramètres ROI ---
$CostPerHour = 30
$MinutesSavedPerJob = 15
$MonthlyRPACost = 5900

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

# --- Folders ---
$FoldersResponse = Invoke-RestMethod -Uri "$BaseUrl/Folders" -Headers $Headers -Method Get
$Folders = $FoldersResponse.value | Sort-Object DisplayName -Unique

# --- Fonction générique pour récupérer les jobs ---
function Get-UipathJobs {
    param (
        [int]$DaysAgo
    )

    $NowUtc = (Get-Date).ToUniversalTime()
    $FilterDate = $NowUtc.AddDays(-$DaysAgo).ToString("yyyy-MM-ddTHH:mm:ssZ")

    function Get-UipathJobsForFolder {
        param ([string]$FolderId, [string]$FolderName)

        $FolderHeaders = $Headers.Clone()
        $FolderHeaders.Add("X-UIPATH-OrganizationUnitId", $FolderId)

        $Jobs = @()
        $NextUrl = "$BaseUrl/Jobs?`$filter=(EndTime ge $FilterDate) and (State eq 'Successful' or State eq 'Faulted' or State eq 'Stopped')&`$orderby=EndTime desc&`$top=100"

        while ($NextUrl) {
            try {
                $Response = Invoke-RestMethod -Uri $NextUrl -Headers $FolderHeaders -Method Get
                $Jobs += $Response.value
                $NextUrl = $Response.'@odata.nextLink'
            } catch { break }
        }

        foreach ($job in $Jobs) {
            $job | Add-Member -NotePropertyName FolderName -NotePropertyValue $FolderName
        }
        return $Jobs
    }

    $AllJobs = @()
    foreach ($folder in $Folders) {
        $AllJobs += Get-UipathJobsForFolder -FolderId $folder.Id -FolderName $folder.DisplayName
    }
    return $AllJobs
}

# --- Création Excel ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
if (Test-Path $XlsPath) { $wb = $excel.Workbooks.Open($XlsPath) } else { $wb = $excel.Workbooks.Add(); $wb.SaveAs($XlsPath, 51) }

# --- Périodes ---
$Periods = @(
    @{Days=30; Name="J-30"},
    @{Days=7; Name="J-7"},
    @{Days=1; Name="J-1"}
)

$Summaries = @{}

foreach ($p in $Periods) {
    $AllJobs = Get-UipathJobs -DaysAgo $p.Days

    # --- Feuille Datas ---
    $DatasName = "Datas_$($p.Name)"
    try { $ws = $wb.Worksheets.Item($DatasName); $ws.Cells.Clear() } catch { $ws = $wb.Worksheets.Add(); $ws.Name = $DatasName }
    $headers = 'Id','ReleaseName','State','StartTime','EndTime','FolderName'
    for ($i=0; $i -lt $headers.Count; $i++) { $ws.Cells.Item(1, $i+1) = $headers[$i] }
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

    # --- Summary ---
    $SummaryName = "Summary_$($p.Name)"
    try { $wsSummary = $wb.Worksheets.Item($SummaryName); $wsSummary.Cells.Clear() } catch { $wsSummary = $wb.Worksheets.Add(); $wsSummary.Name = $SummaryName }
    $headersSummary = 'FolderName','TotalJobs','Successful','Faulted','Stopped','SuccessRate','TotalHoursSaved','ROI','GainNet'
    for ($i=0; $i -lt $headersSummary.Count; $i++) { $wsSummary.Cells.Item(1, $i+1) = $headersSummary[$i] }

    $SummaryData = @()
    foreach ($folder in $Folders) {
        $JobsInFolder = $AllJobs | Where-Object { $_.FolderName -eq $folder.DisplayName }
        if ($JobsInFolder.Count -eq 0) { continue }

        $Success = ($JobsInFolder | Where-Object {$_.State -eq 'Successful'}).Count
        $Faulted = ($JobsInFolder | Where-Object {$_.State -eq 'Faulted'}).Count
        $Stopped = ($JobsInFolder | Where-Object {$_.State -eq 'Stopped'}).Count
        $Total = $Success + $Faulted + $Stopped

        $Taux = if ($Total -gt 0) { [math]::Round(($Success / $Total), 2) } else { $null }
        $TotalMinutesSaved = $Success * $MinutesSavedPerJob
        $TotalHoursSaved = [math]::Round(($TotalMinutesSaved / 60), 2)
        $TotalValue = $TotalHoursSaved * $CostPerHour
        $ROI = if ($MonthlyRPACost -gt 0) { [math]::Round(($TotalValue / $MonthlyRPACost), 2) } else { $null }
        $GainNet = [math]::Round(($TotalValue - $MonthlyRPACost),2)

        $SummaryData += [PSCustomObject]@{
            FolderName      = $folder.DisplayName
            TotalJobs       = $Total
            Successful      = $Success
            Faulted         = $Faulted
            Stopped         = $Stopped
            SuccessRate     = $Taux
            TotalHoursSaved = $TotalHoursSaved
            ROI             = $ROI
            GainNet         = $GainNet
        }
    }

    $Summaries[$p.Name] = $SummaryData

    $row = 2
    foreach ($item in $SummaryData) {
        $wsSummary.Cells.Item($row,1) = $item.FolderName
        $wsSummary.Cells.Item($row,2) = $item.TotalJobs
        $wsSummary.Cells.Item($row,3) = $item.Successful
        $wsSummary.Cells.Item($row,4) = $item.Faulted
        $wsSummary.Cells.Item($row,5) = $item.Stopped
        $wsSummary.Cells.Item($row,6).Value2 = [double]$item.SuccessRate
        $wsSummary.Cells.Item($row,7).Value2 = [double]$item.TotalHoursSaved
        $wsSummary.Cells.Item($row,8).Value2 = [double]$item.ROI
        $wsSummary.Cells.Item($row,9).Value2 = [double]$item.GainNet
        $row++
    }
}

# --- Comparaison J-30 vs J-7 ---
$wsCompare = $wb.Worksheets.Item("Summary_J-30")
$wsCompare.Cells.Item(1,10) = "ΔSuccessRate"
$wsCompare.Cells.Item(1,11) = "ΔROI"
$wsCompare.Cells.Item(1,12) = "ΔGainNet"

foreach ($row in 2..($wsCompare.UsedRange.Rows.Count)) {
    $folderName = $wsCompare.Cells.Item($row,1).Text
    if (-not $folderName) { continue }

    $Data30 = $Summaries["J-30"] | Where-Object { $_.FolderName -eq $folderName }
    $Data7 = $Summaries["J-7"] | Where-Object { $_.FolderName -eq $folderName }

    if ($Data30 -and $Data7) {
        $wsCompare.Cells.Item($row,10).Value2 = [math]::Round(($Data30.SuccessRate - $Data7.SuccessRate),2)
        $wsCompare.Cells.Item($row,11).Value2 = [math]::Round(($Data30.ROI - $Data7.ROI),2)
        $wsCompare.Cells.Item($row,12).Value2 = [math]::Round(($Data30.GainNet - $Data7.GainNet),2)
    }
}

# --- Sauvegarde & nettoyage ---
$wb.Save()
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "✅ Export complet terminé :"
Write-Host "Feuilles Datas_J-30 / J-7 / J-1 + Summary_J-30 / J-7 / J-1 avec comparaison ajoutée."
