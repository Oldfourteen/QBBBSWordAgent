# Run with Windows PowerShell 5.1. No Python, Pester or Office required.
param([switch]$Live)
$ErrorActionPreference='Stop'
. "$PSScriptRoot/../src/native/Environment.ps1"
. "$PSScriptRoot/../src/native/Main.ps1"
. "$PSScriptRoot/../src/native/Maintenance.ps1"
$originalRoot=$script:Root
$testRoot=Join-Path $originalRoot ('tmp/staging/native-tests-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
$manifest=Read-Json (Join-Path $originalRoot 'config/agent_integrity.json')
foreach ($relative in @($manifest.files.Keys)+@('config/agent_integrity.json','README.md')) {
    $dest=Join-Path $testRoot $relative; [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($dest)); [IO.File]::Copy((Join-Path $originalRoot $relative),$dest)
}
$script:Root=$testRoot
$passed=0
function Assert($Condition,$Label) { if (-not $Condition) {throw "FAIL: $Label"}; $script:passed++; Write-Output "PASS: $Label" }
function Must-Fail([scriptblock]$Action,$Label) { $failed=$false; try { & $Action | Out-Null } catch {$failed=$true}; Assert $failed $Label }
Must-Fail {Assert-SupportedHost ([PlatformID]::Unix) ([version]'7.4') $true} 'non-Windows rejected'
Must-Fail {Assert-SupportedHost ([PlatformID]::Win32NT) ([version]'5.0') $true} 'old PowerShell rejected'
Must-Fail {Assert-SupportedHost ([PlatformID]::Win32NT) ([version]'5.1') $false} 'missing worker rejected'
Assert-SupportedHost ([PlatformID]::Win32NT) ([version]'5.1') $true
Assert $true 'Windows PowerShell 5.1 accepted'
Invoke-Native @('start','--topic','原生自动化验证报告','--date','20990101') | Out-Null
$sp=Join-Path $testRoot 'dir/20990101/workflow_state.json'
Assert ((Read-Json $sp).status -eq 'T1_TERM_REQUIRED') 'start requires academic term'
Assert ((Read-Json $sp).input_resources.choice -eq 'empty') 'empty input skips consent and cleanup'
Must-Fail {Invoke-Native @('set-term','--state',$sp,'--academic-year','2025-2027','--semester','1','--confirm')} 'non-consecutive academic year rejected'
Must-Fail {Invoke-Native @('set-term','--state',$sp,'--academic-year','2025-2026','--semester','3','--confirm')} 'invalid semester rejected'
Invoke-Native @('set-term','--state',$sp,'--academic-year','2025-2026','--semester','2','--confirm') | Out-Null
Must-Fail {Invoke-Native @('choose-code','--state',$sp,'--mode','none','--confirm')} 'cannot bypass gate1'
$outline="# 原生测试大纲`n## 环境能力验证`n- 验证环境。`n### 原生文件构建`n- 验证构建。`n## 文档结构验证`n- 验证结构。`n### 目录页码验证`n- 验证目录。"
$draft=Join-Path ([IO.Path]::GetDirectoryName($sp)) 'outline.draft.md'; [IO.File]::WriteAllText($draft,$outline,$script:Utf8)
$foreign=Join-Path $testRoot 'foreign-outline.md';[IO.File]::WriteAllText($foreign,$outline,$script:Utf8)
Must-Fail {Invoke-Native @('save-outline','--state',$sp,'--source',$foreign)} 'historical/foreign outline path rejected'
Invoke-Native @('save-outline','--state',$sp,'--source',$draft) | Out-Null
Must-Fail {Invoke-Native @('approve-outline','--state',$sp,'--pages','1')} 'gate1 requires confirmation'
Invoke-Native @('approve-outline','--state',$sp,'--pages','1','--confirm') | Out-Null
Assert ((Read-Json $sp).status -eq 'REPORT_TITLE_WAITING_CHOICE') 'title question follows gate1'
Must-Fail {Invoke-Native @('choose-code','--state',$sp,'--mode','none','--confirm')} 'code selection blocked before title'
Must-Fail {Invoke-Native @('set-report-title','--state',$sp,'--mode','default','--project-name','原生测试项目')} 'title requires confirmation'
Invoke-Native @('set-report-title','--state',$sp,'--mode','default','--project-name','原生测试项目','--confirm') | Out-Null
Assert ((Read-Json $sp).report_identity.report_title -eq '基于原生测试项目开发实训') 'default title uses project without duplicated suffix'
Assert ((New-ReportIdentity custom '图书管理系统' '《图书借阅系统设计与实现》').report_title -eq '图书借阅系统设计与实现') 'custom title preserved'
Assert ((Get-IntakeOptions 2026).academic_years -join ',' -eq '2025-2026,2026-2027') 'academic year suggestions use current year'
Invoke-Native @('choose-code','--state',$sp,'--mode','none','--confirm') | Out-Null
Invoke-Native @('prepare-content','--state',$sp) | Out-Null
$input=Join-Path ([IO.Path]::GetDirectoryName($sp)) 'report_content.json'
Must-Fail {Invoke-Native @('prepare-content','--state',$sp)} 'content not overwritten'
Must-Fail {Invoke-Native @('generate','--workflow-state',$sp,'--input',$input,'--output','test.docx','--check-only')} 'placeholder rejected'
$data=Read-Json $input
foreach ($s in $data.sections) { $s.paragraphs=@('用于测试原生文档构建。'); foreach ($sub in $s.subsections) {$sub.paragraphs=@('用于测试二级标题和目录关联。')} }
Write-Json $input $data
Invoke-Native @('generate','--workflow-state',$sp,'--input',$input,'--output','test.docx','--check-only') | Out-Null
Assert ((Read-Json $sp).status -eq 'T2_READY') 'precheck does not advance gate'
$pending=Join-Path $testRoot 'test.pending.docx'; Build-Document $data $pending $null
$zip=[IO.Compression.ZipFile]::OpenRead($pending)
try { $xml=Read-SafeXml (Read-ZipText $zip 'word/document.xml'); $text=$xml.OuterXml; Assert ($text.Contains('1. 环境能力验证') -and $text.Contains('一、环境能力验证')) 'distinct TOC/body numbering'; Assert ($text.Contains('1.1 原生文件构建')) 'subheadings present' } finally {$zip.Dispose()}
Assert ($text.Contains('《基于原生测试项目开发实训》实践周成绩报告单') -and $text.Contains('《基于原生测试项目开发实训》实践周总结') -and $text.Contains('2025-2026学年度第二学期') -and $text.Contains('2025—2026学年第二学期期末考试')) 'both front pages share confirmed title and term'
$realTemplate=Join-Path $originalRoot 'samples/school_templates/成绩单、封面和撰写要求样本1.docx'
if(Test-Path $realTemplate){
    $originalTemplateHash=Hash $realTemplate
    $customData=Read-Json $input;$customData.report_title='图书借阅系统设计与实现';$customData.project_name='图书管理系统';$customData.front_matter=@{report_title=$customData.report_title;project_name=$customData.project_name;academic_year='2026-2027';semester=1}
    $customDoc=Join-Path $testRoot 'custom.pending.docx';Build-Document $customData $customDoc $realTemplate
    $customZip=[IO.Compression.ZipFile]::OpenRead($customDoc)
    try{$customXml=Read-SafeXml (Read-ZipText $customZip 'word/document.xml');$customText=$customXml.DocumentElement.InnerText}finally{$customZip.Dispose()}
    Assert ($customText.Contains('《图书借阅系统设计与实现》实践周成绩报告单') -and $customText.Contains('《图书借阅系统设计与实现》实践周总结') -and $customText.Contains('2026-2027学年度第一学期') -and $customText.Contains('2026—2027学年第一学期期末考试') -and -not $customText.Contains('《基于xxx项目开发实训》') -and -not $customText.Contains('《基于XXXXXXXXX开发实训》')) 'real template custom title and first semester replace split placeholders'
    Assert ((Hash $realTemplate) -eq $originalTemplateHash) 'source school template preserved'
}
$map=@{status='ok';body_pages=1;pages=5;toc_start_physical_page=4;toc_end_physical_page=4;body_start_physical_page=5;front_matter_pages=@{_QBScoreTitle=1;_QBScoreTerm=1;_QBCoverTitle=2;_QBCoverTerm=2};page_map=@()}
foreach ($h in (Get-Headings $data)) {$map.page_map+=@{bookmark=$h.bookmark;toc_page=1;body_footer_page=1}}
Must-Fail {Test-SavedDocument $pending $map $data} 'unrefreshed cached page rejected'
if ($Live) {
    $liveDir=Join-Path $testRoot 'live'; [void][IO.Directory]::CreateDirectory($liveDir)
    $liveData=Read-Json (Join-Path $originalRoot 'dir/20260930/report_content.json')
    $liveData.report_title=$data.report_title;$liveData.project_name=$data.project_name;$liveData.front_matter=$data.front_matter
    $template=Join-Path $originalRoot 'samples/school_templates/成绩单、封面和撰写要求样本1.docx'
    if (-not (Test-Path $template)) {$template=$null}
    $liveDoc=Join-Path $liveDir 'test.pending.docx'; Build-Document $liveData $liveDoc $template
    $engines=Get-Renderers; if (-not $engines.Count) {throw 'Live test requires a registered engine.'}
    $liveMap=Invoke-Renderer $liveDoc (Join-Path $liveDir 'diagnostics') $engines[0] 90
    Test-SavedDocument $liveDoc $liveMap $liveData
    Assert ($liveMap.front_matter_pages.Count -eq 4) 'live front matter fields located on physical pages 1 and 2'
    Assert ($liveMap.body_pages -gt 1) 'live multipage template TOC/footer and fonts'
    Write-Output ($liveMap | ConvertTo-Json -Depth 8 -Compress)
}
# Mock rendering only in this isolated test copy; never touch live task receipts.
function Select-Renderer($State,$StatePath,$Requested) { return @{engine='test-mock';prog_id='test'} }
function Invoke-Renderer($Document,$Directory,$Engine,$Timeout) {
    $z=[IO.Compression.ZipFile]::Open($Document,[IO.Compression.ZipArchiveMode]::Update)
    try { $body=Read-ZipText $z 'word/document.xml'; $body=$body.Replace('>0</w:t>','>1</w:t>'); Write-ZipText $z 'word/document.xml' $body } finally {$z.Dispose()}
    return $map
}
Invoke-Native @('generate','--workflow-state',$sp,'--input',$input,'--output','test.docx') | Out-Null
Assert ((Read-Json $sp).status -eq 'GATE3_WAITING_APPROVAL') 'mock render enters gate3 only'
Must-Fail {Invoke-Native @('finish','--state',$sp,'--output','final.docx','--confirm')} 'cannot bypass gate3/4'
Invoke-Native @('generate','--workflow-state',$sp,'--input',$input,'--output','test.docx') | Out-Null
Assert ((Read-Json $sp).status -eq 'GATE3_WAITING_APPROVAL') 'verified reuse retains gate'
Invoke-Native @('review-decision','--state',$sp,'--approve','--confirm') | Out-Null
Invoke-Native @('final-check','--state',$sp) | Out-Null
Must-Fail {Invoke-Native @('finish','--state',$sp,'--output','final.docx')} 'gate4 requires confirmation'
Invoke-Native @('finish','--state',$sp,'--output','final.docx','--confirm') | Out-Null
Assert ((Read-Json $sp).status -eq 'COMPLETE') 'mock full workflow completion'
$state=Read-Json $sp; $receipt=Project-Path $state.review.receipt; $tampered=Read-Json $receipt; $tampered.body_pages=99; Write-Json $receipt $tampered
Must-Fail {Check-Review $state $sp} 'receipt tampering rejected'
Must-Fail {Project-Path '../outside.docx'} 'path escape rejected'
Assert ((Get-InputInventory).Count -eq 0) 'missing input directory is empty'
$fixture=Join-Path $testRoot 'input/project';[void][IO.Directory]::CreateDirectory($fixture)
$sourceFile=Join-Path $fixture 'Main.java';[IO.File]::WriteAllText($sourceFile,'class Main {}',$script:Utf8)
Invoke-Native @('start','--topic','材料入口验证','--date','20990102') | Out-Null
$isp=Join-Path $testRoot 'dir/20990102/workflow_state.json'
$is=Read-Json $isp
Assert ($is.input_resources.choice -eq 'pending' -and $is.input_resources.hashes.Count -eq 0) 'nonempty input inventories metadata without reading contents'
Must-Fail {Invoke-Native @('start','--topic','不可混批','--date','20990103')} 'pending input batch blocks another task'
Must-Fail {Invoke-Native @('set-term','--state',$isp,'--academic-year','2025-2026','--semester','1','--confirm')} 'term cannot bypass input consent'
Must-Fail {Invoke-Native @('choose-input','--state',$isp,'--mode','read')} 'input reading requires explicit confirmation'
Invoke-Native @('choose-input','--state',$isp,'--mode','read','--confirm') | Out-Null
$is=Read-Json $isp
Assert ($is.input_resources.choice -eq 'read' -and $is.input_resources.hashes.Count -eq 1) 'read confirmation binds resource hashes'
Assert-InputReady $is $isp
[IO.File]::WriteAllText($sourceFile,'class Changed {}',$script:Utf8)
Must-Fail {Assert-InputReady $is $isp} 'changed materials cannot silently enter generation'
Must-Fail {Invoke-Native @('clean','--input','--state',$isp)} 'unfinished report cannot clean input'
Must-Fail {Invoke-Native @('clean','--input','--retired','--state',$isp)} 'input cleanup cannot mix unrelated modes'
# Simulate publication only inside isolated test state; no real report approvals changed.
$is.status='COMPLETE';Write-Json $isp $is
Must-Fail {Invoke-Native @('clean','--input','--state',$isp)} 'automatic cleanup refuses changed input batch'
Must-Fail {Invoke-Native @('clean','--input','--state',$isp,'--manual')} 'manual move must actually leave input empty'
$moved=Join-Path $testRoot 'manually-moved-project';[IO.Directory]::Move($fixture,$moved)
Invoke-Native @('clean','--input','--state',$isp,'--manual') | Out-Null
Must-Fail {Invoke-Native @('clean','--input','--state',$isp,'--manual','--apply')} 'manual cleanup finalization requires confirmation'
Invoke-Native @('clean','--input','--state',$isp,'--manual','--apply','--confirm') | Out-Null
Assert ((Read-Json $isp).input_resources.cleanup -eq 'completed') 'manual move releases batch after empty verification'
Assert (Test-Path $moved) 'manually moved project preserved'
[void][IO.Directory]::CreateDirectory($fixture);[IO.File]::WriteAllText($sourceFile,'class Main {}',$script:Utf8)
Invoke-Native @('start','--topic','跳过读取验证','--date','20990103') | Out-Null
$ssp=Join-Path $testRoot 'dir/20990103/workflow_state.json'
Invoke-Native @('choose-input','--state',$ssp,'--mode','skip','--confirm') | Out-Null
$ss=Read-Json $ssp
Assert ($ss.input_resources.choice -eq 'skip' -and $ss.input_resources.hashes.Count -eq 0) 'skip records no content reading'
$ss.status='COMPLETE';Write-Json $ssp $ss
Must-Fail {Invoke-Native @('clean','--input','--state',$ssp,'--apply','--confirm')} 'input apply requires preview'
Invoke-Native @('clean','--input','--state',$ssp) | Out-Null
Assert (Test-Path $sourceFile) 'cleanup preview does not delete resources'
Must-Fail {Invoke-Native @('clean','--input','--state',$ssp,'--apply')} 'automatic cleanup requires confirmation'
$newFile=Join-Path $fixture 'new-material.txt';[IO.File]::WriteAllText($newFile,'preserve new material',$script:Utf8)
Must-Fail {Invoke-Native @('clean','--input','--state',$ssp,'--apply','--confirm')} 'new materials after preview block deletion'
Assert (Test-Path $newFile) 'blocked cleanup preserves newly added material'
[IO.File]::Delete($newFile) # Only remove this isolated test fixture.
# Recycle operation mocked only for isolated fixtures, never user's actual input.
function Move-InputItemToRecycleBin($Target) {
    if (-not $Target.StartsWith((Join-Path $testRoot 'input')+'\')) {throw 'Test removal outside fixture refused.'}
    if (Test-Path -LiteralPath $Target -PathType Container) {[IO.Directory]::Delete($Target,$true)} else {[IO.File]::Delete($Target)}
}
Invoke-Native @('clean','--input','--state',$ssp,'--apply','--confirm') | Out-Null
Assert ((Get-InputInventory).Count -eq 0 -and (Test-Path (Join-Path $testRoot 'input'))) 'input cleanup empties batch but preserves root'
Assert ((Read-Json $ssp).input_resources.cleanup -eq 'completed') 'successful cleanup recorded and owner released'
Assert (Test-Path $moved) 'input cleanup preserves materials outside input'
Invoke-Native @('start','--topic','下一批验证','--date','20990104') | Out-Null
Assert ((Read-Json (Join-Path $testRoot 'dir/20990104/workflow_state.json')).input_resources.choice -eq 'empty') 'next empty batch starts without permission question'
Write-Output "Native tests passed: $passed. Workflow Office calls mocked; optional Live=$Live used real Office. Artifacts: $testRoot"
