. "$PSScriptRoot/Core.ps1"
. "$PSScriptRoot/Document.ps1"
. "$PSScriptRoot/Runtime.ps1"
. "$PSScriptRoot/Intake.ps1"
. "$PSScriptRoot/Input.ps1"
function Check-Review($State,$StatePath) {
    if (-not $State.review.native) { throw 'Legacy review needs native regeneration before approval; no silent trust in old receipts.' }
    $doc=Project-Path $State.review.document; $receipt=Project-Path $State.review.receipt
    if ((Hash $doc) -ne $State.review.sha256 -or (Hash $receipt) -ne $State.review.receipt_sha256) { throw 'Review artifact changed.' }
    $r=Read-Json $receipt
    if ($r.sha256 -ne (Hash $doc) -or $r.body_pages -ne $State.target_body_pages -or $r.code_block_mode -ne $State.code_block_mode) { throw 'Review receipt differs from approval.' }
    if ($r.implementation_sha256 -ne (Hash (Join-Path $script:Root 'config/agent_integrity.json'))) { throw 'Implementation changed since generation; regenerate review.' }
    [void](Approved-Outline $State $StatePath)
    Assert-TaskFrontMatter $r.content $State
    Test-SavedDocument $doc $r $r.content
    return $r
}
function Invoke-Native($Arguments) {
    Assert-ReportEnvironment
    Check-Integrity
    $script:Policy=Read-Json (Join-Path $script:Root 'config/report_policy.json')
    if (-not $Arguments.Count) { throw 'Command required. Read workflow/COMMANDS.md.' }
    $command=$Arguments[0]; $o=@{}
    for ($i=1;$i -lt $Arguments.Count;$i++) {
        $key=[string]$Arguments[$i]
        if (-not $key.StartsWith('--')) { throw "Unexpected argument: $key" }
        $key=$key.Substring(2); if ($o.ContainsKey($key)) { throw "Duplicate option: $key" }
        if ($i+1 -lt $Arguments.Count -and -not ([string]$Arguments[$i+1]).StartsWith('--')) { $i++; $o[$key]=$Arguments[$i] } else { $o[$key]=$true }
    }
    $allowed=@{
        workflow=@(); 'input-status'=@(); 'choose-input'=@('state','mode','confirm'); 'intake-options'=@(); doctor=@('probe-word','probe','engine'); start=@('topic','date'); status=@('state'); 'set-term'=@('state','academic-year','semester','confirm'); 'set-report-title'=@('state','mode','project-name','title','confirm'); 'save-outline'=@('state','source'); 'approve-outline'=@('state','pages','confirm'); 'choose-code'=@('state','mode','confirm'); 'prepare-content'=@('state'); generate=@('workflow-state','input','output','school-template','engine','word-timeout','check-only','force-rebuild'); 'review-decision'=@('state','approve','reject','confirm'); 'final-check'=@('state'); finish=@('state','output','confirm'); rollback=@('state','to','reason'); package=@(); clean=@('apply','retired','runtime','scratch','input','state','manual','confirm'); 'render-job'=@('job')
    }
    if (-not $allowed.ContainsKey($command)) { throw "Unknown command: $command" }
    foreach ($key in $o.Keys) { if ($allowed[$command] -notcontains $key) { throw "Unknown option: --$key" } }
    if ($command -in @('start','generate','final-check','finish') -and -not $o.'check-only') { Invoke-RuntimeHygiene 'before_task' }
    if ($command -eq 'workflow') { Get-Content (Join-Path $script:Root 'workflow/WORKFLOW.md') -Raw -Encoding UTF8; return }
    if ($command -eq 'intake-options') {Get-IntakeOptions | ConvertTo-Json -Depth 8;return}
    if ($command -eq 'input-status') { $items=Get-InputInventory; @{empty=($items.Count -eq 0);items=$items;notice='input must be empty before adding a new batch; nonempty current materials require read/skip consent.'} | ConvertTo-Json -Depth 10;return }
    if ($command -eq 'doctor') {
        $envInfo=@{runtime='Windows PowerShell/.NET; Python not required';renderers=(Get-Renderers);runtime_verified=$false;TEMP=$env:TEMP;TMP=$env:TMP;cache=(Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders').Cache}
        if ($o.probe -or $o.'probe-word') {
            $probeDir=Reserve-Directory (Join-Path $script:Root 'tmp/staging') ('doctor-'+[guid]::NewGuid().ToString('N'))
            $probeState=@{history=@()}; $engine=if($o.engine){$o.engine}else{'auto'}
            $envInfo.renderer=Select-Renderer $probeState (Join-Path $probeDir 'probe.json') $engine; $envInfo.runtime_verified=$true
        }
        $envInfo | ConvertTo-Json -Depth 20; return
    }
    if ($command -eq 'package') { New-NativePackage; return }
    if ($command -eq 'clean') { Invoke-NativeClean $o; return }
    if ($command -eq 'render-job') { throw 'Legacy queued jobs remain preserved. Resume using generate with the original workflow-state/input; native generation executes Office automatically. Do not approve gates or manually edit the queued job.' }
    if ($command -eq 'start') {
        if (-not $o.topic -or -not ([string]$o.topic).Trim()) { throw 'Topic required.' }
        $date=if ($o.date) {[string]$o.date}else{Get-Date -Format yyyyMMdd}
        if ($date -notmatch '^\d{8}$') { throw 'Date must be YYYYMMDD.' }
        $inputLock=Open-InputOperationLock
        try {
        Assert-InputAvailable
        $inputItems=Get-InputInventory
        $dir=Reserve-Directory (Join-Path $script:Root 'dir') $date; $statePath=Join-Path $dir 'workflow_state.json'
        $s=@{agent_name='QBWordAgent';workflow_version='4-gate-v2-intake';workflow_sha256=(Hash (Join-Path $script:Root 'workflow/WORKFLOW.md'));workflow_id=[IO.Path]::GetFileName($dir);topic=$o.topic;status='T1_TERM_REQUIRED';academic_term=$null;report_identity=$null;outline_contract_sha256=(Hash (Join-Path $script:Root 'templates/OUTLINE_RULES.md'));target_body_pages=$null;code_block_mode=$null;outline=$null;review=$null;final=$null;gates=@{};history=@();created_at=[DateTimeOffset]::Now.ToString('o')}
        foreach ($g in @('gate1','gate2','gate3','gate4')) {$s.gates[$g]=@{status='closed'}}
        $s.input_resources=New-InputSession $inputItems $statePath
        Save-State $statePath $s 'workflow_created'
        $template=Get-Content (Join-Path $script:Root 'templates/outline.md') -Raw -Encoding UTF8
        [IO.File]::WriteAllText((Join-Path $dir 'outline.draft.md'),$template.Replace('{{TOPIC}}',$o.topic),$script:Utf8)
        Get-Content (Join-Path $script:Root 'workflow/WORKFLOW.md') -Raw -Encoding UTF8
        @{agent_name='QBWordAgent';state=$statePath;status=$s.status;input_resources=$s.input_resources;options=(Get-IntakeOptions);next_action=(Get-TaskNextAction $s)} | ConvertTo-Json -Depth 12; return
        } finally {$inputLock.Dispose()}
    }
    $stateArg=if ($command -eq 'generate') {$o.'workflow-state'}else{$o.state}
    $statePath=Project-Path $stateArg; $s=Load-State $statePath; $dir=[IO.Path]::GetDirectoryName($statePath)
    # Exclusive handle serializes all task mutations; no stale-file deletion is necessary.
    $lockPath=Join-Path $dir 'native-operation.lock'; $lock=$null
    try {
        $lock=[IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        $s=Load-State $statePath
        switch ($command) {
            status { $s | ConvertTo-Json -Depth 40; return }
            'choose-input' { Set-InputChoice $s $statePath $o }
            'set-term' {
                Assert-InputReady $s $statePath
                Require-State $s @('T1_TERM_REQUIRED','T1_OUTLINE_REQUIRED');Require-Confirm $o
                Test-AcademicTerm ([string]$o.'academic-year') $o.semester
                $s.academic_term=@{academic_year=[string]$o.'academic-year';semester=[int]$o.semester;confirmed_at=[DateTimeOffset]::Now.ToString('o')};$s.status='T1_OUTLINE_REQUIRED'
            }
            'save-outline' {
                Assert-InputReady $s $statePath
                Require-State $s @('T1_OUTLINE_REQUIRED','GATE1_WAITING_APPROVAL')
                Require-AcademicTerm $s
                Assert-OutlineContract $s
                $source=Project-Path $o.source
                if([IO.Path]::GetDirectoryName($source) -ne $dir){throw '大纲只能来自本任务目录，不得读取历史任务大纲。'}
                $text=Get-Content $source -Raw -Encoding UTF8; [void](Parse-Outline $text)
                $p=Join-Path $dir 'outline.md'; [IO.File]::WriteAllText($p,$text,$script:Utf8)
                $s.outline=@{path=$p;sha256=(Hash $p)}; $s.status='GATE1_WAITING_APPROVAL'; $s.gates.gate1=@{status='waiting_user'}
            }
            'approve-outline' {
                Require-State $s @('GATE1_WAITING_APPROVAL'); Require-Confirm $o; [void](Approved-Outline $s $statePath)
                Require-AcademicTerm $s
                $pages=[int]$o.pages; if ($pages -lt $script:Policy.limits.min_body_pages -or $pages -gt $script:Policy.limits.max_body_pages) { throw 'Page target outside limits.' }
                $s.target_body_pages=$pages; $s.gates.gate1=@{status='approved';at=[DateTimeOffset]::Now.ToString('o')}; $s.gates.gate2=@{status='closed'}; $s.status='REPORT_TITLE_WAITING_CHOICE'
            }
            'set-report-title' {
                Require-State $s @('REPORT_TITLE_WAITING_CHOICE');Require-Confirm $o;Require-AcademicTerm $s
                $s.report_identity=New-ReportIdentity $o.mode $o.'project-name' $o.title
                $s.status='GATE2_WAITING_CHOICE';$s.gates.gate2=@{status='waiting_user'}
            }
            'choose-code' {
                Require-State $s @('GATE2_WAITING_CHOICE'); Require-Confirm $o
                [void](Get-FrontMatter $s)
                if (@('none','embedded','after_body') -notcontains $o.mode) { throw 'Invalid code mode.' }
                $s.code_block_mode=$o.mode; $s.gates.gate2=@{status='approved';at=[DateTimeOffset]::Now.ToString('o')}; $s.status='T2_READY'
            }
            'prepare-content' {
                Assert-InputReady $s $statePath
                Require-State $s @('T2_READY','T2_REVISION_REQUIRED'); $p=Join-Path $dir 'report_content.json'
                if (Test-Path $p) { throw 'Content already exists; refusing overwrite.' }
                $metadata=Get-FrontMatter $s
                $data=@{report_title=$metadata.report_title;project_name=$metadata.project_name;front_matter=$metadata;code_block_mode=$s.code_block_mode;sections=(Parse-Outline (Get-Content (Approved-Outline $s $statePath) -Raw -Encoding UTF8))}
                if ($s.code_block_mode -ne 'none') { $data.code_blocks=@(@{content='待填写代码';language='';section_index=1}) }
                Write-Json $p $data; Write-Output $p; return
            }
            generate { Invoke-Generate $s $statePath $o; return }
            'review-decision' {
                Require-State $s @('GATE3_WAITING_APPROVAL'); Require-Confirm $o
                if ([bool]$o.approve -eq [bool]$o.reject) { throw 'Choose exactly one of --approve/--reject.' }
                if ($o.approve) { [void](Check-Review $s $statePath); $s.gates.gate3=@{status='approved'}; $s.status='T3_READY' }
                else { $s.gates.gate3=@{status='rejected'}; $s.status='T2_REVISION_REQUIRED' }
            }
            'final-check' { Require-State $s @('T3_READY'); [void](Check-Review $s $statePath); $s.status='GATE4_WAITING_APPROVAL'; $s.gates.gate4=@{status='waiting_user'} }
            finish {
                Assert-InputReady $s $statePath
                Require-State $s @('GATE4_WAITING_APPROVAL'); Require-Confirm $o; $receipt=Check-Review $s $statePath
                $name=[IO.Path]::GetFileName([string]$o.output); if ([IO.Path]::GetExtension($name) -ne '.docx') { throw 'Final output must be .docx.' }
                $out=Reserve-Directory (Join-Path $script:Root 'output') (Get-Date -Format yyyyMMdd); $dest=Join-Path $out $name
                [IO.File]::Copy((Project-Path $s.review.document),$dest,$false)
                $receipt.publication_status='final'; $receipt.document=$name; $receipt.gate4_approved=$true
                Write-Json (Join-Path $out 'verification_receipt.json') $receipt
                Invoke-RuntimeHygiene 'before_complete'
                $s.final=@{document=$dest;sha256=(Hash $dest)}; $s.status='COMPLETE'; $s.gates.gate4=@{status='approved'}
            }
            rollback {
                Require-State $s @('T2_READY','T2_REVISION_REQUIRED','GATE2_WAITING_CHOICE','REPORT_TITLE_WAITING_CHOICE','GATE1_WAITING_APPROVAL','GATE3_WAITING_APPROVAL','T3_READY','GATE4_WAITING_APPROVAL')
                if (-not $o.reason) { throw 'Rollback reason required.' }
                $s.gates.gate3=@{status='closed'}; $s.gates.gate4=@{status='closed'}
                switch ($o.to) {
                    t1 { $s.status='T1_TERM_REQUIRED'; $s.academic_term=$null;$s.report_identity=$null;$s.outline_contract_sha256=Hash (Join-Path $script:Root 'templates/OUTLINE_RULES.md'); $s.outline=$null; $s.target_body_pages=$null; $s.code_block_mode=$null; $s.gates.gate1=@{status='closed'}; $s.gates.gate2=@{status='closed'} }
                    gate1 { $s.status='GATE1_WAITING_APPROVAL';$s.report_identity=$null; $s.target_body_pages=$null; $s.code_block_mode=$null; $s.gates.gate1=@{status='waiting_user'}; $s.gates.gate2=@{status='closed'} }
                    gate2 { $s.status='GATE2_WAITING_CHOICE'; $s.code_block_mode=$null; $s.gates.gate2=@{status='waiting_user'} }
                    t2 { $s.status='T2_REVISION_REQUIRED' }
                    default { throw 'Invalid rollback target.' }
                }
            }
        }
        Save-State $statePath $s "$command $($o.reason)"
        @{agent_name='QBWordAgent';status=$s.status;state=$statePath;next_action=(Get-TaskNextAction $s);message='本步完成，停止并向用户展示结果；按固定工作流等待下一步授权。'} | ConvertTo-Json
    } finally { if ($lock) {$lock.Dispose()} }
}
function Invoke-Generate($State,$StatePath,$Options) {
    $dir=[IO.Path]::GetDirectoryName($StatePath); $watch=[Diagnostics.Stopwatch]::StartNew()
    try {
        Require-State $State @('T2_READY','T2_REVISION_REQUIRED','GATE3_WAITING_APPROVAL')
        Assert-InputReady $State $StatePath
        $inputPath=Project-Path $Options.input
        if([IO.Path]::GetDirectoryName($inputPath) -ne $dir){throw '正文输入只能来自本任务目录，不得套用其他任务JSON。'}
        $data=Read-Json $inputPath; Assert-TaskFrontMatter $data $State; Test-Content $data $State $StatePath
        $name=[IO.Path]::GetFileName([string]$Options.output); if ([IO.Path]::GetExtension($name) -ne '.docx') { throw 'Review output must be .docx.' }
        if ($Options.'check-only') { Write-Output '输入预检通过；没有启动Office，未改变闸门。'; return }
        $template=if ($Options.'school-template') { Project-Path $Options.'school-template' } else { Join-Path $script:Root 'samples/school_templates/成绩单、封面和撰写要求样本1.docx' }
        if ($Options.'school-template' -and -not (Test-Path $template)) { throw 'Requested template not found.' }
        if (-not (Test-Path $template)) { $template=$null }
        $quality=@{template=if($template){Hash $template}else{'built-in'};nodes=@()}
        $nodes=Get-ChildItem (Join-Path $script:Root 'samples/training/rejected/error_correction_constraints') -Filter constraint_node.json -Recurse -ErrorAction SilentlyContinue
        foreach ($node in $nodes) { $quality.nodes+=@{path=$node.FullName.Substring($script:Root.Length+1);sha256=(Hash $node.FullName)} }
        $keyText=($data | ConvertTo-Json -Depth 50 -Compress)+(Hash (Join-Path $script:Root 'config/agent_integrity.json'))+($quality | ConvertTo-Json -Depth 20 -Compress)+$State.target_body_pages
        $sha=[Security.Cryptography.SHA256]::Create(); try { $key=[BitConverter]::ToString($sha.ComputeHash($script:Utf8.GetBytes($keyText))).Replace('-','') } finally {$sha.Dispose()}
        if (-not $Options.'force-rebuild' -and $State.review.native -and $State.review.key -eq $key) {
            [void](Check-Review $State $StatePath)
            if ($State.status -ne 'GATE3_WAITING_APPROVAL') { $State.status='GATE3_WAITING_APPROVAL'; $State.gates.gate3=@{status='waiting_user'}; Save-State $StatePath $State 'review_reused' }
            Write-Output "复用已验证文档：$($State.review.document)。停止等待闸门3。"; return
        }
        Require-State $State @('T2_READY','T2_REVISION_REQUIRED')
        $requested=if($Options.engine){$Options.engine}else{'auto'}
        if (@('auto','word','wps') -notcontains $requested) { throw 'Engine must be auto/word/wps.' }
        $engine=Select-Renderer $State $StatePath $requested
        $timeout=if($Options.'word-timeout'){[int]$Options.'word-timeout'}else{180}; if($timeout -lt 10 -or $timeout -gt 600){throw 'Timeout must be 10-600 seconds.'}
        $stage=Reserve-Directory (Join-Path $script:Root 'tmp/staging') ('native-'+[guid]::NewGuid().ToString('N'))
        $pending=Join-Path $stage 'report.pending.docx'; Build-Document $data $pending $template
        $map=Invoke-Renderer $pending (Join-Path $stage 'diagnostics') $engine $timeout
        if ($map.body_pages -ne $State.target_body_pages) { throw "实际正文 $($map.body_pages) 页，用户批准 $($State.target_body_pages) 页；修改正文后重试，不用空段落凑页。" }
        Test-SavedDocument $pending $map $data
        $review=Reserve-Directory (Join-Path $dir 'reviews') 'draft'; $dest=Join-Path $review $name
        [IO.File]::Copy($pending,$dest,$false)
        $map.sha256=Hash $dest; $map.document=$name; $map.content=$data; $map.code_block_mode=$State.code_block_mode; $map.publication_status='t2_review'; $map.renderer=$engine; $map.quality_sources=$quality; $map.implementation_sha256=Hash (Join-Path $script:Root 'config/agent_integrity.json')
        $receipt=Join-Path $review 'verification_receipt.json'; Write-Json $receipt $map
        $State.review=@{native=$true;document=$dest.Substring($script:Root.Length+1);receipt=$receipt.Substring($script:Root.Length+1);sha256=(Hash $dest);receipt_sha256=(Hash $receipt);key=$key}
        $State.status='GATE3_WAITING_APPROVAL'; $State.gates.gate3=@{status='waiting_user'}; $State.gates.gate4=@{status='closed'}
        Invoke-NativeClean @{scratch=$pending} | Out-Null
        Invoke-NativeClean @{scratch=$pending;apply=$true} | Out-Null
        Save-State $StatePath $State 'native_t2_verified'
        Write-Json (Join-Path $dir 'last_generation.json') @{status='generated';elapsed_seconds=$watch.Elapsed.TotalSeconds;output=$dest;renderer=$engine;body_pages=$map.body_pages}
        Write-Output "待审文档：$dest。停止等待用户确认闸门3。"
    } catch {
        Write-Json (Join-Path $dir 'last_generation.json') @{status='failed';elapsed_seconds=$watch.Elapsed.TotalSeconds;error=$_.Exception.Message}
        throw
    }
}
