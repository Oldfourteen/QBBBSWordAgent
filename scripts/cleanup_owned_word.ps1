param([Parameter(Mandatory=$true)][string]$TracePath)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false)
$trace=Get-Content -LiteralPath $TracePath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $trace.owned -or -not $trace.word_pid -or -not $trace.started_ticks) {
    Write-Output 'No proven owned Word process; no process terminated.'
    exit 0
}
$target=Get-Process -Id ([int]$trace.word_pid) -ErrorAction SilentlyContinue
if (-not $target) { Write-Output 'Owned Word process already exited.'; exit 0 }
if ($target.ProcessName -notin @('WINWORD','wps') -or ($trace.process_name -and $target.ProcessName -ne $trace.process_name) -or [string]$target.StartTime.ToUniversalTime().Ticks -ne [string]$trace.started_ticks) {
    throw 'Process identity changed; termination refused.'
}
if ($target.MainWindowTitle) { throw 'Word has a visible document window; termination refused.' }
Stop-Process -Id $target.Id -Force -ErrorAction Stop
Write-Output "Stopped only owned Word process $($target.Id)."
