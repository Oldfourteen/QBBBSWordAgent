param([Parameter(Mandatory=$true)][string]$DocumentPath)
$ErrorActionPreference = "Stop"
$wordApp = $null
$wordDoc = $null
function Step($label, [scriptblock]$block) {
    try { & $block; Write-Output ("OK   : " + $label) }
    catch {
        Write-Output ("FAIL : " + $label + " -> " + $_.Exception.Message)
        throw
    }
}
try {
    Step "create app" { $script:wordApp = New-Object -ComObject Word.Application }
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    Step "open doc" { $script:wordDoc = $wordApp.Documents.Open($DocumentPath, $false, $false, $false) }
    Step "repaginate" { $wordDoc.Repaginate() }
    Step "fields.update" { $wordDoc.Fields.Update() | Out-Null }
    Step "footers loop" {
        foreach ($section in $wordDoc.Sections) {
            foreach ($footer in $section.Footers) {
                if ($footer.Exists) { $footer.Range.Fields.Update() | Out-Null }
            }
        }
    }
    Step "repaginate2" { $wordDoc.Repaginate() }
    Step "fields.update2" { $wordDoc.Fields.Update() | Out-Null }
    Step "compute pages" { $global:_pages = [int]$wordDoc.ComputeStatistics(2) }
    Write-Output ("pages = " + $global:_pages)
    $script:bodySection = $wordDoc.Sections.Item($wordDoc.Sections.Count)
    Write-Output ("sections count = " + $wordDoc.Sections.Count)
    Step "body footer pages" { $global:_bodyPages = [int]$bodySection.Range.ComputeStatistics(2) }
    Write-Output ("body pages = " + $global:_bodyPages)
    $script:primaryFooter = $bodySection.Footers.Item(1)
    Write-Output ("footer exists = " + $primaryFooter.Exists)
    Step "iterate pageref fields" {
        $script:count = 0
        foreach ($field in $wordDoc.Fields) {
            $code = $field.Code.Text.Trim()
            if ($code -notmatch "PAGEREF\s+([^\s\\]+)") { continue }
            $script:count += 1
            $bn = $Matches[1].Trim('"')
            $has = $wordDoc.Bookmarks.Exists($bn)
            Write-Output ("  field#" + $count + " bookmark=" + $bn + " exists=" + $has)
            if (-not $has) { continue }
            $res = $field.Result.Text
            Write-Output ("    result text=[" + $res + "]")
            $r1 = [int]$wordDoc.Bookmarks.Item($bn).Range.Information(1)
            $r3 = [int]$wordDoc.Bookmarks.Item($bn).Range.Information(3)
            Write-Output ("    range info(1)=" + $r1 + " info(3)=" + $r3)
        }
        Write-Output ("  total PAGEREF fields = " + $script:count)
    }
    Write-Output "DIAG_REFRESH_OK"
}
catch {
    Write-Output ("FAILED overall: " + $_.Exception.Message)
}
finally {
    if ($wordDoc) { try { $wordDoc.Close(0) } catch {} }
    if ($wordApp) { try { $wordApp.Quit() } catch {} }
}
