# Reusable input lifecycle. Resource contents are never executed by these tools.
function Get-InputInventory {
    $root=Project-Path 'input'
    $items=[Collections.Generic.List[object]]::new(); $queue=[Collections.Generic.Queue[string]]::new()
    if (Test-Path -LiteralPath $root) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) {throw 'input must be a directory.'}
        $queue.Enqueue($root)
    }
    while ($queue.Count) {
        $parent=$queue.Dequeue()
        foreach ($entry in (Get-ChildItem -LiteralPath $parent -Force -ErrorAction Stop)) {
            $path=Project-Path $entry.FullName # Reject links before descending.
            $relative=$path.Substring($root.Length+1).Replace('\','/')
            $items.Add(@{path=$relative;kind=if($entry.PSIsContainer){'directory'}else{'file'};bytes=if($entry.PSIsContainer){0L}else{$entry.Length};modified_ticks=if($entry.PSIsContainer){0L}else{$entry.LastWriteTimeUtc.Ticks}})
            if ($entry.PSIsContainer) {$queue.Enqueue($path)}
        }
    }
    return ,@($items | Sort-Object { $_.path })
}
function Get-InputSignature($Items) {
    $text=@($Items | ForEach-Object {'{0}|{1}|{2}|{3}' -f $_.path,$_.kind,$_.bytes,$_.modified_ticks}) -join "`n"
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return [BitConverter]::ToString($sha.ComputeHash($script:Utf8.GetBytes($text))).Replace('-','')} finally {$sha.Dispose()}
}
function Get-InputOwnerPath {Join-Path $script:Root 'qa/maintenance/input-owner.json'}
function Open-InputOperationLock {
    $path=Join-Path $script:Root 'qa/maintenance/input-operation.lock'
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
    return [IO.File]::Open($path,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
}
function Assert-InputAvailable {
    $p=Get-InputOwnerPath
    if (Test-Path -LiteralPath $p) {
        $owner=Read-Json $p
        if ($owner.active) {throw "input belongs to an unfinished resource lifecycle. Resume task $($owner.state); move/clean its resources before a new task."}
    }
}
function New-InputSession($Items,$StatePath) {
    $session=@{items=@($Items);signature=(Get-InputSignature $Items);choice=if($Items.Count){'pending'}else{'empty'};cleanup=if($Items.Count){'pending'}else{'not_required'};hashes=@{}}
    if ($Items.Count) {
        $ownerPath=Get-InputOwnerPath;[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($ownerPath))
        Write-Json $ownerPath @{active=$true;state=$StatePath.Substring($script:Root.Length+1);at=[DateTimeOffset]::Now.ToString('o')}
    }
    return $session
}
function Assert-InputOwned($StatePath) {
    $owner=Read-Json (Get-InputOwnerPath)
    if (-not $owner.active -or (Project-Path $owner.state) -ne $StatePath) {throw 'input owner does not match this task.'}
}
function Assert-InputReady($State,$StatePath) {
    if (-not $State.input_resources) {return} # Explicitly preserve legacy tasks.
    $r=$State.input_resources
    if ($r.choice -eq 'pending') {throw 'Ask whether to read input resources and record choose-input before proceeding.'}
    if ($r.choice -eq 'empty') {
        if ((Get-InputInventory).Count) {throw 'input changed after task start; restart intake instead of silently using new materials.'}
        return
    }
    Assert-InputOwned $StatePath
    if ((Get-InputSignature (Get-InputInventory)) -ne $r.signature) {throw 'input resources changed. Restore the approved batch or start a new task; do not silently overwrite the snapshot.'}
    if ($r.choice -eq 'read') {
        foreach ($relative in $r.hashes.Keys) {if ((Hash (Project-Path ('input/'+$relative))) -ne $r.hashes[$relative]) {throw 'Approved input contents changed.'}}
    }
}
function Set-InputChoice($State,$StatePath,$Options) {
    Require-Confirm $Options; Require-State $State @('T1_TERM_REQUIRED')
    $r=$State.input_resources
    if (-not $r -or $r.choice -ne 'pending') {throw 'No pending input choice.'}
    if (@('read','skip') -notcontains $Options.mode) {throw 'Input mode must be read or skip.'}
    Assert-InputOwned $StatePath
    if ((Get-InputSignature (Get-InputInventory)) -ne $r.signature) {throw 'input changed before confirmation; do not consume a different batch.'}
    if ($Options.mode -eq 'read') {
        foreach ($item in $r.items) {if ($item.kind -eq 'file') {$r.hashes[$item.path]=Hash (Project-Path ('input/'+$item.path))}}
    }
    $r.choice=$Options.mode;$r.confirmed_at=[DateTimeOffset]::Now.ToString('o')
}
function Get-TaskNextAction($State) {
    if ($State.input_resources.choice -eq 'pending') {return 'input has materials. Show filenames only and ask read/skip; do not read contents before explicit permission.'}
    if ($State.status -eq 'COMPLETE' -and $State.input_resources.cleanup -eq 'pending') {return 'Report published. Ask whether the user will move input materials manually; otherwise confirm recycling via clean --input preview/apply. input must be empty before the next batch.'}
    return Get-NextIntakeAction $State.status
}
function Invoke-InputClean($Options) {
    $statePath=Project-Path $Options.state;$lockPath=Join-Path ([IO.Path]::GetDirectoryName($statePath)) 'native-operation.lock'
    $inputLock=Open-InputOperationLock;$lock=$null
    try {
        $lock=[IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        $state=Load-State $statePath;Require-State $state @('COMPLETE')
        if (-not $state.input_resources -or $state.input_resources.cleanup -ne 'pending') {throw 'No pending input cleanup for this task.'}
        Assert-InputOwned $statePath
        $root=Project-Path 'input';$items=Get-InputInventory
        $mode=if($Options.manual){'manual'}else{'recycle-bin'}
        $planPath=Join-Path ([IO.Path]::GetDirectoryName($statePath)) 'input-cleanup.json'
        if (-not $Options.apply) {
            if ($Options.manual) {if ($items.Count) {throw 'Manual move is not complete; input must be empty.'}}
            else {Assert-InputReady $state $statePath}
            $files=@();foreach($item in $items){if($item.kind -eq 'file'){$files+=@{path=$item.path;sha256=(Hash (Project-Path ('input/'+$item.path)))}}}
            $tops=@(Get-ChildItem -LiteralPath $root -Force -ErrorAction SilentlyContinue | ForEach-Object {$_.Name})
            Write-Json $planPath @{status='preview';mode=$mode;signature=(Get-InputSignature $items);files=$files;top_level=$tops;removed=@()}
            @{plan=$planPath;mode=$mode;items=$items;notice='Only this task input batch. Recycle Bin recovery; input directory itself is preserved. Explicit consent required for apply.'} | ConvertTo-Json -Depth 10
            return
        }
        Require-Confirm $Options;$plan=Read-Json $planPath
        if ($plan.status -ne 'preview' -or $plan.mode -ne $mode -or $plan.signature -ne (Get-InputSignature $items)) {throw 'Cleanup needs an unchanged preview with the same mode.'}
        foreach($file in $plan.files){if ((Hash (Project-Path ('input/'+$file.path))) -ne $file.sha256) {throw 'Cleanup file contents changed after preview.'}}
        if ($mode -eq 'recycle-bin') {
            if ((Get-InputSignature $items) -ne $state.input_resources.signature) {throw 'Cleanup batch differs from task materials.'}
            Add-Type -AssemblyName Microsoft.VisualBasic
            $plan.status='applying';Write-Json $planPath $plan
            try {
                foreach($name in $plan.top_level){
                    $target=Project-Path ('input/'+$name)
                    if ([IO.Path]::GetDirectoryName($target) -ne $root) {throw 'Cleanup must stay directly within input.'}
                    Move-InputItemToRecycleBin $target
                    $plan.removed+= $name;Write-Json $planPath $plan
                }
            } catch {$plan.status='partial';$plan.error=$_.Exception.Message;Write-Json $planPath $plan;throw}
        }
        if ((Get-InputInventory).Count) {throw 'input is not empty; lifecycle remains pending.'}
        $plan.status='completed';Write-Json $planPath $plan
        $state.input_resources.cleanup='completed';$state.input_resources.cleanup_mode=$mode;Save-State $statePath $state 'input_empty_verified'
        Write-Json (Get-InputOwnerPath) @{active=$false;state=$statePath.Substring($script:Root.Length+1);at=[DateTimeOffset]::Now.ToString('o')}
        @{status='completed';input_empty=$true;removed=$plan.removed;recovery=if($mode -eq 'recycle-bin'){'Windows Recycle Bin'}else{'User moved materials'};message='input is empty and ready for the next batch.'} | ConvertTo-Json -Depth 8
    } finally {if($lock){$lock.Dispose()};$inputLock.Dispose()}
}
function Move-InputItemToRecycleBin($Target) {
    if (Test-Path -LiteralPath $Target -PathType Container) {[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($Target,[Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,[Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin,[Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)}
    else {[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($Target,[Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,[Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin,[Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)}
}
