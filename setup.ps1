# Compatibility entry: no dependency installation or system changes.
& (Join-Path $PSScriptRoot 'agent.ps1') doctor
exit $LASTEXITCODE
