function New-NativePackage {
    $manifest=Read-Json (Join-Path $script:Root 'config/agent_integrity.json')
    $files=@($manifest.files.Keys)+@('config/agent_integrity.json','README.md','LICENSE','tests/test_native.ps1')
    $dist=Join-Path $script:Root 'dist'; [void][IO.Directory]::CreateDirectory($dist)
    $path=Join-Path $dist ('QBWordAgent-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid().ToString('N').Substring(0,6)+'.zip')
    $stream=[IO.File]::Open($path,[IO.FileMode]::CreateNew)
    $zip=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Create)
    try {
        [void]$zip.CreateEntry('QBWordAgent/input/')
        foreach ($file in ($files | Sort-Object -Unique)) {
            $entry=$zip.CreateEntry('QBWordAgent/'+$file.Replace('\','/')); $src=[IO.File]::OpenRead((Project-Path $file)); $dst=$entry.Open()
            try {$src.CopyTo($dst)}finally{$src.Dispose();$dst.Dispose()}
        }
        $nodes=Get-ChildItem (Join-Path $script:Root 'samples/training/rejected/error_correction_constraints') -Filter constraint_node.json -Recurse -ErrorAction SilentlyContinue
        foreach ($node in $nodes) {
            $source=Read-Json $node.FullName; $safe=@{}
            foreach ($key in @('node_id','status','severity','rules','enforcement_layers')) { if ($source.ContainsKey($key)) {$safe[$key]=$source[$key]} }
            $entry=$zip.CreateEntry('QBWordAgent/'+$node.FullName.Substring($script:Root.Length+1).Replace('\','/')); $writer=[IO.StreamWriter]::new($entry.Open(),$script:Utf8)
            try {$writer.Write(($safe | ConvertTo-Json -Depth 40))}finally{$writer.Dispose()}
        }
    } finally {$zip.Dispose();$stream.Dispose()}
    @{package=$path;sha256=(Hash $path);note='No historical reports, user samples, Python runtime or Office included.'} | ConvertTo-Json
}
function Invoke-NativeClean($Options) {
    if ($Options.input) {
        if ($Options.runtime -or $Options.scratch -or $Options.retired) {throw 'Input cleanup cannot be combined with other cleanup modes.'}
        Invoke-InputClean $Options;return
    }
    if ($Options.state -or $Options.manual -or $Options.confirm) {throw 'state/manual/confirm require clean --input.'}
    if ($Options.runtime) { Invoke-RuntimeHygiene 'explicit_cleanup'; return }
    if ($Options.scratch) { Invoke-ScratchClean $Options; return }
    if ($Options.retired) { Invoke-RetiredClean $Options; return }
    $root=Join-Path $script:Root 'tmp/staging'; $files=@()
    if (Test-Path $root) {
        $files=@(Get-ChildItem $root -File -Recurse | Where-Object { $_.Name -in @('stdout.txt','stderr.txt') -and $_.LastWriteTime -lt (Get-Date).AddDays(-14) })
    }
    $plan=@($files | ForEach-Object {@{path=$_.FullName;bytes=$_.Length;sha256=(Hash $_.FullName)}})
    if ($Options.apply) {
        foreach ($item in $plan) { $p=Project-Path $item.path; if (-not $p.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Cleanup outside staging refused.' }; if ((Hash $p) -ne $item.sha256) {throw 'Cleanup target changed.'}; Remove-Item -LiteralPath $p }
    }
    @{applied=[bool]$Options.apply;plan=$plan;note='Only diagnostic stdout/stderr older than 14 days; reports, samples and task data preserved.'} | ConvertTo-Json -Depth 10
}
function Invoke-ScratchClean($Options) {
    $path=Project-Path $Options.scratch
    $relative=$path.Substring($script:Root.Length+1).Replace('\','/')
    if ($relative -notmatch '^tmp/staging/native-(?:probe-)?[a-f0-9]{32}/(?:report|Word\.Application|KWPS\.Application|WPS\.Application)\.pending\.docx$') {throw 'Scratch cleanup outside generated staging refused.'}
    $planPath=Join-Path ([IO.Path]::GetDirectoryName($path)) 'scratch-cleanup.json'
    if (-not $Options.apply) { Write-Json $planPath @{path=$relative;sha256=(Hash $path);status='preview'}; return }
    $plan=Read-Json $planPath
    if ($plan.status -ne 'preview' -or $plan.path -ne $relative -or (Hash $path) -ne $plan.sha256) {throw 'Scratch cleanup requires unchanged preview.'}
    Remove-Item -LiteralPath $path -ErrorAction Stop
    $plan.status='removed_verified_scratch'; Write-Json $planPath $plan
}
function Invoke-RetiredClean($Options) {
    $planPath=Join-Path $script:Root 'qa/maintenance/retired-cleanup-plan.json'
    $approved=Read-Json (Join-Path $script:Root 'config/retired_cleanup.json')
    if (-not $Options.apply) {
        $items=@()
        foreach ($relative in $approved.files) {
            $path=Project-Path $relative
            if (Test-Path -LiteralPath $path -PathType Leaf) { $items+=@{path=$relative;sha256=(Hash $path);bytes=(Get-Item -LiteralPath $path).Length} }
        }
        Write-Json $planPath @{files=$items;created_at=[DateTimeOffset]::Now.ToString('o');mode='recycle-bin';status='preview'}
        $totalBytes=0L; foreach($item in $items){$totalBytes+=$item.bytes}
        @{plan=$planPath;count=$items.Count;bytes=$totalBytes;mode='recycle-bin'} | ConvertTo-Json
        return
    }
    $plan=Read-Json $planPath
    if ($plan.status -ne 'preview') {throw 'Run clean --retired preview first.'}
    $protected=(Read-Json (Join-Path $script:Root 'config/workflow_policy.json')).protected_paths
    foreach ($item in $plan.files) {
        if ($approved.files -notcontains $item.path -or $protected -contains $item.path -or $item.path -match '^(dir|output|samples)[/\\]') {throw 'Cleanup target not authorized.'}
        $path=Project-Path $item.path
        if ((Hash $path) -ne $item.sha256) {throw "Cleanup target changed: $path"}
    }
    Add-Type -AssemblyName Microsoft.VisualBasic
    $plan.status='applying'; $plan.removed=@(); Write-Json $planPath $plan
    try {
        foreach ($item in $plan.files) {
            $path=Project-Path $item.path
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($path,[Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,[Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin,[Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)
            $plan.removed+= $item.path
        }
        $plan.status='completed'
    } catch { $plan.status='partial'; $plan.error=$_.Exception.Message; throw }
    finally { Write-Json $planPath $plan }
    $totalBytes=0L; foreach($item in $plan.files){$totalBytes+=$item.bytes}
    @{status=$plan.status;count=$plan.removed.Count;bytes=$totalBytes;recovery='Windows Recycle Bin';receipt=$planPath} | ConvertTo-Json
}
