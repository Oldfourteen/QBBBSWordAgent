# No filesystem writes or Windows-only calls before the platform gate.
function Assert-SupportedHost($Platform, $Version, $WorkerExists) {
    if ($Platform -ne [PlatformID]::Win32NT) { throw '任务已取消：仅支持 Windows 10/11，不创建报告任务。' }
    if ([version]$Version -lt [version]'5.1' -or -not $WorkerExists) { throw '任务已取消：需要可运行的 Windows PowerShell 5.1。' }
}
function Assert-ReportEnvironment {
    $platform=[Environment]::OSVersion.Platform
    if ($platform -ne [PlatformID]::Win32NT) { Assert-SupportedHost $platform $PSVersionTable.PSVersion $false }
    $worker=Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
    Assert-SupportedHost $platform $PSVersionTable.PSVersion (Test-Path -LiteralPath $worker -PathType Leaf)
    $windows=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    if ([int]$windows.CurrentBuildNumber -lt 10240 -or $windows.InstallationType -ne 'Client') { throw '任务已取消：本生成器仅支持 Windows 10/11 桌面系统。' }
}
