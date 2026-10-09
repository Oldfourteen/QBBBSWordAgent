function Get-Renderers {
    $list=@()
    foreach ($item in @(@('word','Word.Application'),@('wps','KWPS.Application'),@('wps','WPS.Application'))) {
        if ([type]::GetTypeFromProgID($item[1])) { $list+=@{engine=$item[0];prog_id=$item[1]} }
    }
    return ,$list
}
function Invoke-Renderer($Document,$Directory,$Engine,$Timeout=90) {
    Invoke-RuntimeHygiene 'before_render'
    [void][IO.Directory]::CreateDirectory($Directory)
    $trace=Join-Path $Directory 'trace.json'; $stdout=Join-Path $Directory 'stdout.txt'; $stderr=Join-Path $Directory 'stderr.txt'
    $worker=Join-Path $script:Root 'scripts/refresh_word_fields.ps1'
    $command=@('-NoProfile','-NonInteractive','-File',('"'+$worker+'"'),'-DocumentPath',('"'+$Document+'"'),'-TracePath',('"'+$trace+'"'),'-ProgId',$Engine.prog_id)
    $process=Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList $command -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    $processHandle=$process.Handle # Keep the native handle alive so PS5.1 retains ExitCode.
    $finished=$process.WaitForExit($Timeout*1000)
    if (-not $finished) { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue }
    $process.Refresh()
    if (-not $finished -or $process.ExitCode -ne 0) {
        $detail=if (Test-Path $stderr) { Get-Content $stderr -Raw -Encoding UTF8 } else { '' }
        $stage=if (Test-Path $trace) { (Read-Json $trace).stage } else { 'worker_start' }
        if (Test-Path $trace) {
            $identity=Read-Json $trace
            if ($identity.owned) {
                try { & (Join-Path $script:Root 'scripts/cleanup_owned_word.ps1') -TracePath $trace | Out-Null }
                catch { $detail+="`nOwned-process cleanup refused: $($_.Exception.Message)" }
            }
        }
        throw "Renderer failed ($($Engine.engine)), stage=$stage, timeout=$(-not $finished). $detail Diagnostics: $Directory. Request the host client's legitimate execution approval if access was denied; do not ask the user to copy terminal commands or retry blindly."
    }
    $lines=@(Get-Content $stdout -Encoding UTF8 | Where-Object {$_.Trim()})
    if (-not $lines.Count) { throw 'Renderer did not return a page map.' }
    $identity=Read-Json $trace
    if (-not $identity.owned -or $identity.stage -ne 'completed') {throw '排版未完成安全退出，不能进入下一步。'}
    $remaining=Get-Process -Id ([int]$identity.word_pid) -ErrorAction SilentlyContinue
    if ($remaining -and [string]$remaining.StartTime.ToUniversalTime().Ticks -eq [string]$identity.started_ticks) {
        & (Join-Path $script:Root 'scripts/cleanup_owned_word.ps1') -TracePath $trace | Out-Null
        if (-not $remaining.WaitForExit(5000)) {throw '排版进程清理失败，停止后续生成。'}
    }
    Invoke-RuntimeHygiene 'after_render'
    return (ConvertTo-Map ($lines[-1] | ConvertFrom-Json))
}
function Select-Renderer($State,$StatePath,$Requested='auto') {
    $available=Get-Renderers
    if ($State.renderer) {
        if ($Requested -ne 'auto' -and $Requested -ne $State.renderer.engine) { throw 'Renderer pinned for this task; do not silently change pagination engines.' }
        if (-not @($available | Where-Object {$_.prog_id -eq $State.renderer.prog_id}).Count) { throw 'Pinned renderer unavailable.' }
        return $State.renderer
    }
    if ($Requested -ne 'auto') { $available=@($available | Where-Object {$_.engine -eq $Requested}) }
    if (-not $available.Count) { throw 'No registered Word/WPS automation engine. Installation alone does not prove automation support.' }
    $folder=Reserve-Directory (Join-Path $script:Root 'tmp/staging') ('native-probe-'+[guid]::NewGuid().ToString('N'))
    $data=@{report_title='排版能力验证报告';code_block_mode='none';sections=@(@{title='环境功能检查';paragraphs=@('验证打开与保存。');subsections=@(@{title='页码功能检查';paragraphs=@('验证目录与页码。')})},@{title='排版结果检查';paragraphs=@('验证正文结构。');subsections=@(@{title='最终保存检查';paragraphs=@('验证最终结果。')})})}
    $failures=@()
    foreach ($candidate in $available) {
        $path=Join-Path $folder ($candidate.prog_id+'.pending.docx')
        Build-Document $data $path $null
        try {
            $map=Invoke-Renderer $path (Join-Path $folder $candidate.prog_id) $candidate 45
            Test-SavedDocument $path $map $data
            Invoke-NativeClean @{scratch=$path} | Out-Null
            Invoke-NativeClean @{scratch=$path;apply=$true} | Out-Null
            $candidate.verified_at=[DateTimeOffset]::Now.ToString('o')
            $State.renderer=$candidate; Save-State $StatePath $State 'renderer_capability_verified'
            return $candidate
        } catch {
            $failures+= $_.Exception.Message
            # Stop on the first failed live probe: a hung desktop process must be diagnosed before launching another engine.
            break
        }
    }
    throw ($failures -join "`n")
}
function Invoke-RuntimeHygiene($Phase) {
    $mutex=[Threading.Mutex]::new($false,'Local\FinalReportWordAutomation'); $locked=$false
    $removed=@(); $checked=0
    try {
        try {$locked=$mutex.WaitOne(0)} catch [Threading.AbandonedMutexException] {$locked=$true}
        if (-not $locked) {throw '另一个Office排版任务正在运行；不清理活动进程，请稍后继续。'}
        $staging=Join-Path $script:Root 'tmp/staging'
        if (Test-Path $staging) {
            $folders=Get-ChildItem -LiteralPath $staging -Directory | Where-Object {$_.Name -match '^native-(?:probe-)?[a-f0-9]{32}$'}
            foreach ($folder in $folders) {
                [void](Project-Path $folder.FullName)
                foreach ($file in (Get-ChildItem -LiteralPath $folder.FullName -Filter trace.json -Recurse -File)) {
                    $tracePath=Project-Path $file.FullName; $trace=Read-Json $tracePath; $checked++
                    if (-not $trace.owned -or -not $trace.word_pid -or -not $trace.started_ticks) {continue}
                    $target=Get-Process -Id ([int]$trace.word_pid) -ErrorAction SilentlyContinue
                    if (-not $target) {continue}
                    # PID reuse never authorizes terminating a different process.
                    if ([string]$target.StartTime.ToUniversalTime().Ticks -ne [string]$trace.started_ticks) {continue}
                    & (Join-Path $script:Root 'scripts/cleanup_owned_word.ps1') -TracePath $tracePath | Out-Null
                    if (-not $target.WaitForExit(5000)) {throw '本任务Office进程仍未退出，停止后续生成；需要检查权限或占用。'}
                    $removed+= [int]$trace.word_pid
                }
            }
        }
        $receiptRoot=Join-Path $script:Root 'qa/maintenance'; [void][IO.Directory]::CreateDirectory($receiptRoot)
        Write-Json (Join-Path $receiptRoot 'runtime-cleanup-last.json') @{phase=$Phase;status='owned_processes_clear';checked=$checked;removed_pids=$removed;at=[DateTimeOffset]::Now.ToString('o');scope='Only traced, owned Office processes; unknown/user processes are never terminated.'}
    } finally {if($locked){$mutex.ReleaseMutex()};$mutex.Dispose()}
}
