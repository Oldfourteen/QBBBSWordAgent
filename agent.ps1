# Native Windows entry point. Python is not part of the execution path.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
try {
    . (Join-Path $PSScriptRoot 'src/native/Environment.ps1')
    Assert-ReportEnvironment
    . (Join-Path $PSScriptRoot 'src/native/Main.ps1')
    . (Join-Path $PSScriptRoot 'src/native/Maintenance.ps1')
    Invoke-Native $args
    exit 0
} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 2 }
