param([Parameter(Mandatory=$true)][string]$DocumentPath)
$ErrorActionPreference = "Stop"
$wordApp = $null
$wordDoc = $null
try {
    Write-Output "step: create app"
    $wordApp = New-Object -ComObject Word.Application
    Write-Output ("step: app created, version=" + $wordApp.Version)
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    try { $wordApp.AutomationSecurity = 3 } catch { Write-Output "note: could not set AutomationSecurity" }
    Write-Output "step: opening document"
    $wordDoc = $wordApp.Documents.Open($DocumentPath, $false, $false, $false)
    Write-Output "step: document opened"
    $wordDoc.Repaginate()
    Write-Output "step: repaginated"
    $wordDoc.Fields.Update() | Out-Null
    Write-Output "step: fields updated"
    $wordDoc.Repaginate()
    $pages = [int]$wordDoc.ComputeStatistics(2)
    Write-Output ("step: computed pages=" + $pages)
    $wordDoc.Save()
    Write-Output "step: saved"
    Write-Output "DIAG_OK"
}
catch {
    Write-Output ("FAILED at: " + $_.Exception.Message)
    Write-Output ("COM: " + $_.Exception.GetType().FullName)
}
finally {
    if ($wordDoc) { try { $wordDoc.Close(0) } catch {} }
    if ($wordApp) { try { $wordApp.Quit() } catch {} }
}
