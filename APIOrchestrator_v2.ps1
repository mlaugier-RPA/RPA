<# =====================================================================
# API Orchestrator <=> PowerBI - Script Finalisé avec dynamique des départements et des Temps d'exécutions RPA
# Version : v11.0 
# Auteur : Maxime LAUGIER 
# Update du 23/10/2025
#>

# Variables globales du script avec les 2 fichiers sur les Gdrives rpa-adv@extia.fr => Dossier Reporting PowerBI
$Org = "extiavqvkelj"
$Tenant = "DefaultTenant"
$XlsPath = "H:\Mon Drive\Reporting PowerBI\UiPathJobs.xlsx"
$XlsDept_et_Temps = "H:\Mon Drive\Reporting PowerBI\Departement_Tps_Excecutions.xlsx"
$BaseUrl = "https://cloud.uipath.com/$Org/$Tenant/orchestrator_/odata"

$CostPerHour = 30
$MonthlyRPACost = 5900

# Modules d'excel pour booster les exports surtout des logs 
Import-Module ImportExcel -ErrorAction Stop

# === Supprimer de l'ancien fichier Excel en force => Il faut vérifier qu'il n'y a pas de process Excel.Exe en cours ===
if (Test-Path $XlsPath) { Remove-Item $XlsPath -Force }

# Authentification UiPath avec les bon headers ===
Write-Host "🔑 Authentification UiPath..."
$headers = @{ "Content-Type" = "application/x-www-form-urlencoded" }
$body = "client_id=92615cee-13a8-4195-b52a-3543976033cc&client_secret=lOa%5EtVshMA!mLwLsI8kbwNO)8QH%23p1c%23Qa_jmIN%3FCkYo~YOevEs73EVc(Cb(N2jy&grant_type=client_credentials"

try {
    $response = Invoke-RestMethod "https://cloud.uipath.com/$Org/identity_/connect/token" -Method POST -Headers $headers -Body $body
    $PAT = $response.access_token
    if (-not $PAT) { throw "Token introuvable" }
    Write-Host "✅ Jeton récupéré."
} catch {
    Write-Host "❌ Erreur d'authentification : $($_.Exception.Message)" -ForegroundColor Red
    exit
}

$Headers = @{
    "Authorization" = "Bearer $PAT"
    "Accept" = "application/json;odata=nometadata"
}

# Récupération des dossiers dans le cloud Orchestrator
try {
    $Folders = (Invoke-RestMethod -Uri "$BaseUrl/Folders" -Headers $Headers).value
    Write-Host "📁 $(@($Folders).Count) dossiers trouvés."
} catch {
    Write-Host "❌ Erreur récupération dossiers : $($_.Exception.Message)" -ForegroundColor Red
    exit
}


# Chargement dynamique départements / temps d'exécution
Write-Host "📘 Chargement fichier département/temps..."
$deptData = Import-Excel -Path $XlsDept_et_Temps | Where-Object { $_.'FolderName' -and $_.'Departement' }

$DeptMapping = @{}
foreach ($row in $deptData) {
    $key = $row.'FolderName'.Trim()
    $DeptMapping[$key] = [PSCustomObject]@{
        Departement = $row.'Departement'
        TotalHoursSaved = [math]::Round([double]$row.'Temps d''execution avant RPA'/60,2) # Conversion minutes → heures
    }
}

# Fonctions de clean sur l'excel pour repartir sur de la fresh data
function Clean-ExcelValue($val) {
    try {
        if ($null -eq $val) { return "" }
        elseif ($val -is [System.Array]) { $val = ($val -join ", ") }
        elseif ($val -isnot [string]) { $val = [string]$val }
        $val = $val -replace '[\x00-\x1F]', ''
        if ($val.Length -gt 32000) { $val = $val.Substring(0,32000) + " [TRONQUÉ]" }
        return $val
    } catch { return "⚠️ Valeur illisible" }
}

# Fonctions qui permet de récupérer les jobs d'UIpath depuis l'API Orchestrator
function Get-UipathJobsForFolder {
    param([string]$FolderId, [string]$FolderName, [datetime]$StartDate)
    $FolderHeaders = $Headers.Clone()
    $FolderHeaders["X-UIPATH-OrganizationUnitId"] = "$FolderId"

    $Jobs = @()
    $FilterDate = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $NextUrl = "$BaseUrl/Jobs?`$filter=(CreationTime ge $FilterDate)&`$orderby=CreationTime desc&`$top=1000"

    while ($NextUrl) {
        try {
            $Response = Invoke-RestMethod -Uri $NextUrl -Headers $FolderHeaders -Method Get
            if ($Response.value) { $Jobs += $Response.value }
            $NextUrl = $Response.'@odata.nextLink'
        } catch {
            Write-Host "❌ Erreur Jobs $FolderName : $($_.Exception.Message)" -ForegroundColor Red
            $NextUrl = $null
        }
    }

    foreach ($job in $Jobs) {
        $job | Add-Member -NotePropertyName FolderName -NotePropertyValue $FolderName -Force
        if ($DeptMapping.ContainsKey($FolderName)) {
            $job | Add-Member -NotePropertyName Departement -NotePropertyValue $DeptMapping[$FolderName].Departement -Force
            $job | Add-Member -NotePropertyName TotalHoursSaved -NotePropertyValue $DeptMapping[$FolderName].TotalHoursSaved -Force
        } else {
            $job | Add-Member -NotePropertyName Departement -NotePropertyValue "Autre" -Force
            $job | Add-Member -NotePropertyName TotalHoursSaved -NotePropertyValue 0 -Force
        }
    }

    Write-Host "📦 [$FolderName] Jobs : $($Jobs.Count)"
    return $Jobs
}

# Fonctions qui permet de récupérer les logs d'UIpath depuis l'API Orchestrator
function Get-UipathLogsForFolder {
    param([string]$FolderId, [string]$FolderName, [datetime]$StartDate)
    $FolderHeaders = $Headers.Clone()
    $FolderHeaders["X-UIPATH-OrganizationUnitId"] = "$FolderId"

    $Logs = @()
    $FilterDate = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $NextUrl = "$BaseUrl/RobotLogs?`$filter=(TimeStamp ge $FilterDate)&`$orderby=TimeStamp desc&`$top=1000"

    while ($NextUrl) {
        try {
            $Response = Invoke-RestMethod -Uri $NextUrl -Headers $FolderHeaders -Method Get
            if ($Response.value) { $Logs += $Response.value }
            $NextUrl = $Response.'@odata.nextLink'
        } catch {
            Write-Host "❌ Erreur Logs $FolderName : $($_.Exception.Message)" -ForegroundColor Red
            $NextUrl = $null
        }
    }

    foreach ($log in $Logs) {
        $log | Add-Member -NotePropertyName FolderName -NotePropertyValue $FolderName -Force
        $log | Add-Member -NotePropertyName MachineName -NotePropertyValue $log.WindowsIdentity -Force
    }

    Write-Host "📝 [$FolderName] Logs : $($Logs.Count)"
    return $Logs
}

# Fonctions qui permet d'exporter sur l'excel Summary J-24 / J7d et J-30D
function Export-Summary {
    param([array]$AllJobs)
    $AllStates = @("Successful","Faulted","Stopped","Running","Pending","Terminated","Suspended","Waiting","Stopping")
    $TotalSuccessfulAll = ($AllJobs | Where-Object { $_.State -eq "Successful" }).Count
    if ($TotalSuccessfulAll -eq 0) { $TotalSuccessfulAll = 1 }

    $Summary = @()
    $Grouped = $AllJobs | Group-Object FolderName

    foreach ($group in $Grouped) {
        $Folder = $group.Name
        $Jobs = $group.Group
        if ($Jobs.Count -eq 0) { continue }

        $StateCounts = @{}
        foreach ($s in $AllStates) { $StateCounts[$s] = ($Jobs | Where-Object { $_.State -eq $s }).Count }

        $Success = $StateCounts["Successful"]
        $Completed = ($Jobs | Where-Object { $_.State -in @("Successful","Faulted","Stopped","Terminated") }).Count
        $SuccessRate = if ($Completed -gt 0) { [math]::Round($Success/$Completed,2) } else { 0 }

        $Dept = $Jobs[0].Departement
        $TotalHoursSaved = ($Jobs | Measure-Object -Property TotalHoursSaved -Sum).Sum

        $ProportionalCost = $MonthlyRPACost * ($Success / $TotalSuccessfulAll)
        $HumanEquivalentCost = $TotalHoursSaved * $CostPerHour
        $GainNet = [math]::Round($HumanEquivalentCost - $ProportionalCost, 2)
        $ROI = if ($ProportionalCost -ne 0) { [math]::Round($GainNet / $ProportionalCost, 2) } else { 0 }

        $Summary += [PSCustomObject]@{
            FolderName = $Folder
            Departement = $Dept
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


# les 3 fonctions correspondants 3 3 types de sheet que nous allons avoir, pas besoin de -AutoSize -BoldTopRow sinon erreurs
function Export-JobsToSheet { param($AllJobs,$SheetName) $AllJobs | Export-Excel -Path $XlsPath -WorksheetName $SheetName -ClearSheet }
function Export-SummaryToSheet { param($SummaryData,$SheetName) $SummaryData | Export-Excel -Path $XlsPath -WorksheetName $SheetName -ClearSheet }
function Export-LogsToSheet { param($AllLogs,$SheetName) $AllLogs | Export-Excel -Path $XlsPath -WorksheetName $SheetName -ClearSheet }


# Périodes de temps sur laquel je me base pour les exports
$NowUtc = (Get-Date).ToUniversalTime()
$Periods = @{
    "J1"  = $NowUtc.AddHours(-25)
    "J7"  = $NowUtc.AddHours(-(7*24+1))
    "J30" = $NowUtc.AddHours(-(30*24+1))
}

# Boucle principale qui donne les valeurs par rapport au temps
foreach ($period in $Periods.Keys) {
    Write-Host "`n=== Extraction $period ===" -ForegroundColor Cyan
    $AllJobs = @()
    $AllLogs = @()

    foreach ($folder in $Folders) {
        Start-Sleep -Milliseconds 50
        $AllJobs += Get-UipathJobsForFolder -FolderId $folder.Id -FolderName $folder.DisplayName -StartDate $Periods[$period]
        $AllLogs += Get-UipathLogsForFolder -FolderId $folder.Id -FolderName $folder.DisplayName -StartDate $Periods[$period]
    }

    $SummaryData = Export-Summary -AllJobs $AllJobs

    Export-JobsToSheet -AllJobs $AllJobs -SheetName "Datas_$period"
    Export-SummaryToSheet -SummaryData $SummaryData -SheetName "Summary_$period"
    Export-LogsToSheet -AllLogs $AllLogs -SheetName "Logs_$period"
}

# Script OK???????
Write-Host "`n💾 Export terminé : $XlsPath" -ForegroundColor Green
