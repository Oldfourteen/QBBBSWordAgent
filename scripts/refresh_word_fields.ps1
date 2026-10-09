param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,
    [string]$TracePath,
    [ValidateSet('Word.Application','KWPS.Application','WPS.Application')]
    [string]$ProgId = 'Word.Application'
)

$ErrorActionPreference = "Stop"
$resolvedPath = [System.IO.Path]::GetFullPath($DocumentPath)
$wordApp = $null
$wordDoc = $null
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$priorWordIds = @(Get-Process WINWORD,wps -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$trace = [ordered]@{ stage='starting'; word_pid=$null; started_ticks=$null; owned=$false }
function Write-Stage([string]$stage) {
    $trace.stage=$stage
    $trace.at=[DateTime]::UtcNow.ToString('o')
    if ($TracePath) { $trace | ConvertTo-Json -Compress | Set-Content -LiteralPath $TracePath -Encoding UTF8 }
}
$mutex = [System.Threading.Mutex]::new($false, 'Local\FinalReportWordAutomation')
$locked = $false
$completed = $false

try {
    try { $locked = $mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $locked=$true }
    if (-not $locked) { throw 'Another report Word operation is active in this desktop session.' }
    Write-Stage 'creating_word'
    $wordApp = New-Object -ComObject $ProgId
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class ReportWordWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid); }'
    $reportWordPid = [uint32]0
    if ($wordApp.Documents.Count -ne 0) { throw 'Automation has existing documents; refusing to modify the user session.' }
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    $wordApp.AutomationSecurity = 3
    Write-Stage 'opening_document'
    $wordDoc = $wordApp.Documents.Open($resolvedPath, $false, $false, $false)
    # Word exposes Hwnd on Window, not Application (unlike Excel).
    [void][ReportWordWindow]::GetWindowThreadProcessId([IntPtr]$wordDoc.ActiveWindow.Hwnd, [ref]$reportWordPid)
    $reportWordProcess = Get-Process -Id $reportWordPid
    $trace.word_pid = $reportWordPid
    $trace.process_name = $reportWordProcess.ProcessName
    $trace.started_ticks = [string]$reportWordProcess.StartTime.ToUniversalTime().Ticks
    $trace.owned = $priorWordIds -notcontains $reportWordPid
    if (-not $trace.owned) { throw 'Automation used an existing Office process; only the test document will be closed, not the application.' }

    Write-Stage 'pagination_and_fields'
    $wordDoc.Repaginate()
    $wordDoc.Fields.Update() | Out-Null
    foreach ($section in $wordDoc.Sections) {
        foreach ($footer in $section.Footers) {
            if ($footer.Exists) {
                $footer.Range.Fields.Update() | Out-Null
            }
        }
    }
    $wordDoc.Repaginate()
    $wordDoc.Fields.Update() | Out-Null

    Write-Stage 'checking_page_map'
    $errors = New-Object System.Collections.Generic.List[string]
    $pageMap = New-Object System.Collections.Generic.List[object]
    $pageReferences = 0
    $tocPhysicalPages = New-Object System.Collections.Generic.List[int]
    $bodyHeadingPhysicalPages = New-Object System.Collections.Generic.List[int]
    $bodySection = $wordDoc.Sections.Item($wordDoc.Sections.Count)
    $primaryFooter = $bodySection.Footers.Item(1)
    if (-not $primaryFooter.Exists) {
        $errors.Add("Body section has no primary footer.")
    }
    elseif (-not $primaryFooter.PageNumbers.RestartNumberingAtSection) {
        $errors.Add("Body footer page numbering does not restart at the body section.")
    }
    elseif ([int]$primaryFooter.PageNumbers.StartingNumber -ne 1) {
        $errors.Add("Body footer page numbering does not start at 1.")
    }
    $previousPage = 0
    foreach ($field in $wordDoc.Fields) {
        $fieldCode = $field.Code.Text.Trim()
        if ($fieldCode -notmatch "PAGEREF\s+([^\s\\]+)") {
            continue
        }
        $pageReferences += 1
        $tocPhysicalPages.Add([int]$field.Result.Information(3))
        $bookmarkName = $Matches[1].Trim('"')
        if (-not $wordDoc.Bookmarks.Exists($bookmarkName)) {
            $errors.Add("Missing bookmark for TOC reference: $bookmarkName")
            continue
        }
        $expectedPage = [int]$wordDoc.Bookmarks.Item($bookmarkName).Range.Information(1)
        $bodyHeadingPhysicalPages.Add([int]$wordDoc.Bookmarks.Item($bookmarkName).Range.Information(3))
        $actualText = ($field.Result.Text -replace "[^0-9]", "")
        if ([string]::IsNullOrWhiteSpace($actualText)) {
            $errors.Add("TOC reference has no numeric page result: $bookmarkName")
            continue
        }
        $actualPage = [int]$actualText
        if ($expectedPage -lt $previousPage) {
            $errors.Add("Heading pages are not monotonic: $bookmarkName, page=$expectedPage, previous=$previousPage")
        }
        $previousPage = $expectedPage
        if ($actualPage -ne $expectedPage) {
            $errors.Add("TOC page mismatch: $bookmarkName, result=$actualPage, heading=$expectedPage")
        }
        $pageMap.Add([PSCustomObject]@{
            bookmark = $bookmarkName
            toc_page = $actualPage
            body_footer_page = $expectedPage
        })
    }

    if ($pageReferences -eq 0) {
        $errors.Add("No PAGEREF field was found in the document.")
    }
    if ($pageMap.Count -gt 0 -and [int]$pageMap[0].body_footer_page -ne 1) {
        $errors.Add("The first report heading is not on body footer page 1.")
    }
    $tocStartPhysicalPage = if ($tocPhysicalPages.Count -gt 0) { ($tocPhysicalPages | Measure-Object -Minimum).Minimum } else { 0 }
    $tocEndPhysicalPage = if ($tocPhysicalPages.Count -gt 0) { ($tocPhysicalPages | Measure-Object -Maximum).Maximum } else { 0 }
    $bodyStartPhysicalPage = if ($bodyHeadingPhysicalPages.Count -gt 0) { ($bodyHeadingPhysicalPages | Measure-Object -Minimum).Minimum } else { 0 }
    if ([int]$tocStartPhysicalPage -ne 4) {
        $errors.Add("TOC must start on physical page 4, actual=$tocStartPhysicalPage.")
    }
    if ([int]$tocEndPhysicalPage -lt [int]$tocStartPhysicalPage) {
        $errors.Add("TOC physical page range is invalid.")
    }
    if ([int]$bodyStartPhysicalPage -ne ([int]$tocEndPhysicalPage + 1)) {
        $errors.Add("Body must start immediately after the TOC, toc_end=$tocEndPhysicalPage, body_start=$bodyStartPhysicalPage.")
    }
    if ($errors.Count -gt 0) {
        throw ($errors -join [Environment]::NewLine)
    }
    Write-Stage 'saving_document'
    $frontMatterPages=@{}
    foreach($name in @('_QBScoreTitle','_QBScoreTerm','_QBCoverTitle','_QBCoverTerm')){
        if($wordDoc.Bookmarks.Exists($name)){
            $physical=[int]$wordDoc.Bookmarks.Item($name).Range.Information(3)
            $expected=if($name -like '_QBScore*'){1}else{2}
            if($physical -ne $expected){throw "Front matter field on wrong physical page: $name, expected=$expected, actual=$physical"}
            $frontMatterPages[$name]=$physical
        }
    }
    $wordDoc.Save()
    @{
        status = "ok"
        page_references = $pageReferences
        pages = [int]$wordDoc.ComputeStatistics(2)
        body_pages = [int]$bodySection.Range.ComputeStatistics(2)
        toc_start_physical_page = [int]$tocStartPhysicalPage
        toc_end_physical_page = [int]$tocEndPhysicalPage
        body_start_physical_page = [int]$bodyStartPhysicalPage
        page_map = $pageMap
        front_matter_pages = $frontMatterPages
    } | ConvertTo-Json -Compress
    $completed = $true
}
finally {
    if ($wordDoc) {
        if ($completed) { Write-Stage 'closing_document' }
        try { $wordDoc.Close(0) } catch { [Console]::Error.WriteLine("Document cleanup failed: $_") }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordDoc) } catch { }
        $wordDoc = $null
    }
    if ($wordApp) {
        if ($completed) { Write-Stage 'quitting_word' }
        if ($trace.owned) { try { $wordApp.Quit() } catch { [Console]::Error.WriteLine("Word cleanup failed: $_") } }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordApp) } catch { }
        $wordApp = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($completed) { Write-Stage 'completed' }
    if ($locked) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
