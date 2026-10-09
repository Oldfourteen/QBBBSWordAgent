# QBWordAgent: one source of truth for confirmed front-matter fields.
function Get-IntakeOptions($Year=(Get-Date).Year) {
    @{agent_name='QBWordAgent';academic_years=@("$($Year-1)-$Year","$Year-$($Year+1)");semesters=@(@{value=1;label='第一学期'},@{value=2;label='第二学期'});title_modes=@('default','custom')}
}
function Get-NextIntakeAction($Status) {
    switch($Status){
        T1_TERM_REQUIRED {'请用户选择学年及第一/第二学期；先记录set-term，再设计大纲。'}
        T1_OUTLINE_REQUIRED {'读取唯一大纲规范和固定骨架，仅根据本次材料设计大纲。'}
        GATE1_WAITING_APPROVAL {'请用户批准大纲并指定正文页数。'}
        REPORT_TITLE_WAITING_CHOICE {'请用户选择按实际项目生成题目，或自定义题目；展示完整题目并取得确认，之后才能询问代码块。'}
        GATE2_WAITING_CHOICE {'请用户确认代码块方式：不添加、嵌入正文或正文后。'}
        default {'按workflow执行本步，完成后停止等待下一步确认。'}
    }
}
function Test-AcademicTerm($AcademicYear,$Semester) {
    if ($AcademicYear -notmatch '^(\d{4})-(\d{4})$' -or [int]$Matches[2] -ne ([int]$Matches[1]+1)) {throw '学年必须是连续两年，例如2025-2026。'}
    if ([int]$Semester -notin @(1,2)) {throw '请选择第一学期或第二学期。'}
}
function Require-AcademicTerm($State) {
    if (-not $State.academic_term) {throw '请先询问用户并记录学年、学期；旧任务需回退T1后补充，不得从历史任务猜测。'}
    Test-AcademicTerm $State.academic_term.academic_year $State.academic_term.semester
}
function New-ReportIdentity($Mode,$ProjectName,$Title) {
    $project=([string]$ProjectName).Trim()
    if ($project.Length -lt 4 -or $project.Length -gt 40 -or $project -match '[\r\n]|(?i)x{3,}|待填写|\{\{') {throw '实际项目主题须为4至40字符，不能含占位符或换行。'}
    if ($Mode -eq 'default') {
        if ($Title) {throw '默认题目由实际项目主题生成；自定义题目请选择custom。'}
        $base=$project -replace '项目$',''
        $title="基于${base}项目开发实训"
    } elseif ($Mode -eq 'custom') {
        $title=([string]$Title).Trim()
        if ($title.StartsWith('《') -and $title.EndsWith('》')) {$title=$title.Substring(1,$title.Length-2)}
    } else {throw '题目方式只能是default或custom。'}
    if ($title.Length -lt 4 -or $title.Length -gt 60 -or $title -match '[\r\n]|(?i)x{3,}|待填写|\{\{') {throw '报告题目须为4至60字符，不能含占位符或换行。'}
    foreach($term in $script:Policy.banned_identity_terms){if($title.Contains($term) -or $project.Contains($term)){throw '题目或项目主题含禁止的作者身份信息。'}}
    @{mode=$Mode;project_name=$project;report_title=$title;confirmed_at=[DateTimeOffset]::Now.ToString('o')}
}
function Get-FrontMatter($State) {
    Require-AcademicTerm $State
    if (-not $State.report_identity) {throw '闸门1后必须先确认报告题目，再询问代码块；不得沿用旧报告题目。'}
    $id=$State.report_identity
    $checked=New-ReportIdentity $id.mode $id.project_name $(if($id.mode -eq 'custom'){$id.report_title}else{$null})
    if ($checked.report_title -ne $id.report_title) {throw '确认题目与项目主题不一致。'}
    @{academic_year=$State.academic_term.academic_year;semester=[int]$State.academic_term.semester;report_title=$id.report_title;project_name=$id.project_name}
}
function Get-FrontMatterLabels($Metadata) {
    Test-AcademicTerm $Metadata.academic_year $Metadata.semester
    $semester=if([int]$Metadata.semester -eq 1){'第一学期'}else{'第二学期'}
    @{score_title="《$($Metadata.report_title)》实践周成绩报告单";cover_title="《$($Metadata.report_title)》实践周总结";score_term="$($Metadata.academic_year)学年度$semester";cover_term="$($Metadata.academic_year.Replace('-','—'))学年${semester}期末考试"}
}
function Assert-TaskFrontMatter($Data,$State) {
    $metadata=Get-FrontMatter $State
    if ($Data.report_title -ne $metadata.report_title -or $Data.project_name -ne $metadata.project_name) {throw '正文JSON题目/项目名称必须与用户确认一致；请按本次确认填写。'}
    if ($Data.front_matter) {
        foreach($key in $metadata.Keys){if($Data.front_matter[$key] -ne $metadata[$key]){throw "前置页字段与确认不同：$key"}}
    }
    $Data.front_matter=$metadata
}
