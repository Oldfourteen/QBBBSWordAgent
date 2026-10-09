param(
    [string]$InputJson = "samples/content/java_microservices_report.json",
    [string]$OutputFile = "output/Java微服务知识期末大作业报告.docx",
    [int]$MaxAttempts = 4
)

$ErrorActionPreference = "Continue"

function Reset-Word {
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 45
    $count = (Get-Process WINWORD -ErrorAction SilentlyContinue | Measure-Object).Count
    Write-Output ("Reset-Word: remaining WINWORD = " + $count)
}

for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
    Write-Output ("===== attempt $attempt of $MaxAttempts =====")
    Reset-Word
    $outFile = Join-Path $env:TEMP ("run_agent_" + $attempt + ".out.txt")
    $errFile = Join-Path $env:TEMP ("run_agent_" + $attempt + ".err.txt")
    $proc = Start-Process python -ArgumentList @(
        "scripts/run_agent.py",
        "--input", $InputJson,
        "--output", $OutputFile
    ) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outFile -RedirectStandardError $errFile
    $stdout = if (Test-Path $outFile) { Get-Content $outFile -Raw } else { "" }
    $stderr = if (Test-Path $errFile) { Get-Content $errFile -Raw } else { "" }
    Write-Output ("exit code = " + $proc.ExitCode)
    if ($proc.ExitCode -eq 0) {
        Write-Output "SUCCESS:"
        Write-Output $stdout
        break
    } else {
        Write-Output ("FAILED attempt $attempt. stderr tail:")
        if ($stderr) { Write-Output ($stderr.Substring(0, [Math]::Min(1200, $stderr.Length))) }
    }
}
