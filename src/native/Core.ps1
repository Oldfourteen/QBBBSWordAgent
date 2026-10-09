$script:Root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$script:Utf8=[Text.UTF8Encoding]::new($false)
function Read-Json($Path) { ConvertTo-Map (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) }
function ConvertTo-Map($Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [Management.Automation.PSCustomObject]) { $map=@{}; foreach ($p in $Value.PSObject.Properties) { $map[$p.Name]=ConvertTo-Map $p.Value }; return $map }
    if ($Value -is [Array]) { $a=@(); foreach ($v in $Value) { $a+=,(ConvertTo-Map $v) }; return ,$a }
    return $Value
}
function Write-Json($Path,$Value) {
    $temp="$Path.$([guid]::NewGuid().ToString('N')).pending"
    [IO.File]::WriteAllText($temp,($Value | ConvertTo-Json -Depth 80),$script:Utf8)
    if (Test-Path -LiteralPath $Path) { [IO.File]::Replace($temp,$Path,"$Path.previous") } else { [IO.File]::Move($temp,$Path) }
}
function Hash($Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }
function Project-Path($Path) {
    if (-not $Path) { throw 'Required path is empty.' }
    $full=if ([IO.Path]::IsPathRooted($Path)) { [IO.Path]::GetFullPath($Path) } else { [IO.Path]::GetFullPath((Join-Path $script:Root $Path)) }
    if (-not $full.StartsWith($script:Root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Path must stay inside the project.' }
    $walk=$full
    while ($walk -and $walk -ne $script:Root) {
        if (Test-Path -LiteralPath $walk) { if ((Get-Item -LiteralPath $walk -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Linked paths are not accepted.' } }
        $walk=[IO.Path]::GetDirectoryName($walk)
    }
    return $full
}
function Check-Integrity {
    $policy=Read-Json (Join-Path $script:Root 'config/workflow_policy.json'); $manifest=Read-Json (Join-Path $script:Root 'config/agent_integrity.json')
    if (@($policy.protected_paths).Count -ne $manifest.files.Count) { throw 'Integrity manifest file set differs.' }
    foreach ($p in $policy.protected_paths) { if ((Hash (Project-Path $p)) -ne $manifest.files[$p]) { throw "Protected file changed: $p" } }
}
function Reserve-Directory($Parent,$Base) {
    [void][IO.Directory]::CreateDirectory($Parent)
    $n=1; do { $name=if ($n -eq 1) {$Base} else {'{0}_{1:00}' -f $Base,$n}; $path=Join-Path $Parent $name; $n++ } while (Test-Path -LiteralPath $path)
    [void][IO.Directory]::CreateDirectory($path); return $path
}
function Require-State($State,$Allowed) { if ($Allowed -notcontains $State.status) { throw "Gate blocked. Current state: $($State.status)" } }
function Require-Confirm($Options) { if (-not $Options.ContainsKey('confirm')) { throw 'Explicit user confirmation is required; do not invent approval.' } }
function Save-State($Path,$State,$Event) { $State.updated_at=[DateTimeOffset]::Now.ToString('o'); $State.history=@($State.history)+@(@{at=$State.updated_at;event=$Event}); Write-Json $Path $State }
function Load-State($Path) {
    $state=Read-Json $Path
    if (-not $Path.StartsWith((Join-Path $script:Root 'dir')+'\',[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($Path) -ne 'workflow_state.json') { throw 'Invalid workflow state location.' }
    $compatible=(Read-Json (Join-Path $script:Root 'config/workflow_policy.json')).compatible_workflow_hashes
    if ($state.workflow_sha256 -and $state.workflow_sha256 -ne (Hash (Join-Path $script:Root 'workflow/WORKFLOW.md')) -and $compatible -notcontains $state.workflow_sha256) { throw 'Pinned workflow changed.' }
    return $state
}
function Assert-OutlineContract($State) {
    $compatible=(Read-Json (Join-Path $script:Root 'config/workflow_policy.json')).compatible_outline_contract_hashes
    if($State.outline_contract_sha256 -and $State.outline_contract_sha256 -ne (Hash (Join-Path $script:Root 'templates/OUTLINE_RULES.md')) -and $compatible -notcontains $State.outline_contract_sha256){throw '本任务固定大纲规则发生变化，不能静默换用历史方法。'}
}
function Approved-Outline($State,$StatePath) {
    Assert-OutlineContract $State
    # Local outline is authoritative after moving a whole task to another computer.
    $p=Join-Path ([IO.Path]::GetDirectoryName($StatePath)) 'outline.md'
    if (-not $State.outline -or (Hash $p) -ne $State.outline.sha256) { throw 'Approved outline missing or modified.' }
    return $p
}
function Parse-Outline($Text) {
    $sections=@(); $titles=@()
    if ($Text -match '待填写|\{\{TOPIC\}\}' -or $Text -notmatch '(?m)^#\s+.+' -or $Text -notmatch '(?m)^[-*]\s+\S') { throw 'Incomplete outline.' }
    foreach ($line in ($Text -split '\r?\n')) {
        if ($line.Trim() -match '^(#{2,3})\s+(.+)$') {
            $level=$Matches[1].Length; $title=($Matches[2] -replace '^(?:[一二三四五六七八九十]+[、.．]|\d+(?:\.\d+)?[.、．]?\s+)\s*','').Trim()
            if ($titles -contains $title) { throw 'Duplicate outline heading.' }; $titles+= $title
            $node=@{title=$title;paragraphs=@('待填写：正文内容。')}
            if ($level -eq 2) { $node.subsections=@(); $sections+=,$node }
            else { if (-not $sections.Count) { throw 'Subheading before heading.' }; $sections[-1].subsections+=,$node }
        }
    }
    if ($sections.Count -lt 2 -or @($sections | ForEach-Object {$_.subsections}).Count -lt 2) { throw 'At least two headings and two subheadings required.' }
    if($sections.Count -gt $script:Policy.limits.max_level_1_sections -or @($sections | ForEach-Object {$_.subsections}).Count -gt $script:Policy.limits.max_level_2_sections){throw '大纲标题数量超出固定限制。'}
    foreach ($term in $script:Policy.banned_identity_terms) { if ($Text.Contains($term)) { throw "Identity text prohibited: $term" } }
    return ,$sections
}
function Test-SchemaValue($Value,$Schema,$Label='content') {
    if ($Schema.type) {
        $ok=switch ($Schema.type) {
            object { $Value -is [System.Collections.IDictionary] }
            array { $Value -is [Array] }
            string { $Value -is [string] }
            boolean { $Value -is [bool] }
            integer { $Value -is [int] -or $Value -is [long] }
        }
        if (-not $ok) { throw "Schema type: $Label" }
    }
    if ($Schema.enum -and $Schema.enum -notcontains $Value) { throw "Schema enum: $Label" }
    if ($Value -is [string]) { if (($Schema.minLength -and $Value.Length -lt $Schema.minLength) -or ($Schema.maxLength -and $Value.Length -gt $Schema.maxLength)) { throw "Schema length: $Label" } }
    if ($Value -is [Array]) {
        if (($Schema.minItems -and $Value.Count -lt $Schema.minItems) -or ($Schema.maxItems -and $Value.Count -gt $Schema.maxItems)) { throw "Schema count: $Label" }
        foreach ($v in $Value) { if ($Schema.items) { Test-SchemaValue $v $Schema.items "$Label[]" } }
    }
    if ($Schema.minimum -and $Value -lt $Schema.minimum) { throw "Schema minimum: $Label" }
    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in $Schema.required) { if (-not $Value.ContainsKey($key)) { throw "Missing $Label.$key" } }
        foreach ($key in $Value.Keys) {
            if ($Schema.properties.ContainsKey($key)) { Test-SchemaValue $Value[$key] $Schema.properties[$key] "$Label.$key" }
            else { throw "Unknown field $Label.$key" }
        }
    }
}
function Test-Content($Data,$State,$StatePath) {
    Test-SchemaValue $Data (Read-Json (Join-Path $script:Root 'samples/content/report.schema.json'))
    if (-not $Data.ContainsKey('code_blocks')) { $Data.code_blocks=@() }
    foreach ($section in $Data.sections) { if (-not $section.ContainsKey('subsections')) { $section.subsections=@() } }
    if ($Data.include_appendix -or $Data.appendices) { throw 'Native version does not accept appendix content; default is no appendix.' }
    $text=$Data | ConvertTo-Json -Depth 40
    if ($text -match '待填写|\{\{TOPIC\}\}') { throw 'Unfilled content placeholder.' }
    foreach ($term in $script:Policy.banned_identity_terms) { if ($text.Contains($term)) { throw "Identity prohibited: $term" } }
    if ($Data.code_block_mode -ne $State.code_block_mode) { throw 'Gate 2 code mode differs.' }
    $plan=Parse-Outline (Get-Content -LiteralPath (Approved-Outline $State $StatePath) -Raw -Encoding UTF8)
    $planned=@($plan | ForEach-Object {$_.title; foreach ($s in $_.subsections) {$s.title}})
    $actual=@($Data.sections | ForEach-Object {$_.title; foreach ($s in $_.subsections) {$s.title}})
    if (($planned -join "`n") -ne ($actual -join "`n") -or $plan.Count -ne $Data.sections.Count) { throw 'Headings differ from approved outline.' }
    $subCount=0; $imageCount=0
    for ($i=0;$i -lt $plan.Count;$i++) {
        if (@($plan[$i].subsections).Count -ne @($Data.sections[$i].subsections).Count) { throw 'Subheading parent differs.' }
        foreach ($node in @($Data.sections[$i])+@($Data.sections[$i].subsections)) {
            foreach ($p in $node.paragraphs) { if (-not $p.Trim() -or $p -match '[\r\n]') { throw 'Use nonempty single paragraphs, no blank-line padding.' } }
            if ($node.image_placeholder) { $imageCount++ }
        }
        $subCount+=@($Data.sections[$i].subsections).Count
    }
    if ($subCount -gt $script:Policy.limits.max_level_2_sections -or $imageCount -gt $script:Policy.limits.max_images) { throw 'Content limits exceeded.' }
    if ($Data.code_block_mode -eq 'none' -and @($Data.code_blocks).Count) { throw 'Code blocks prohibited by gate 2.' }
    if ($Data.code_block_mode -ne 'none' -and -not @($Data.code_blocks).Count) { throw 'Code blocks required by gate 2.' }
    foreach ($block in $Data.code_blocks) {
        if (-not $block.content.Trim()) { throw 'Empty code block.' }
        if ($Data.code_block_mode -eq 'embedded') {
            if ($block.section_index -lt 1 -or $block.section_index -gt $Data.sections.Count) { throw 'Invalid code section.' }
            if ($block.subsection_index -and $block.subsection_index -gt @($Data.sections[$block.section_index-1].subsections).Count) { throw 'Invalid code subsection.' }
        }
    }
}
