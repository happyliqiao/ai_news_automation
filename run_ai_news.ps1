param(
    [ValidateSet("all", "ai", "general")]
    [string]$Report = "all"
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = "C:\Users\liqiao\AppData\Local\Programs\Python\Python313\python.exe"
$OutputDir = "E:\AI"
$LogDir = Join-Path $OutputDir "logs"

if (-not (Test-Path -LiteralPath $Python)) {
    throw "找不到 Python：$Python"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$Stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $LogDir "run_$Stamp.log"
$StdoutFile = Join-Path $LogDir "run_$Stamp.stdout.log"
$StderrFile = Join-Path $LogDir "run_$Stamp.stderr.log"

Push-Location $ScriptDir
try {
    $startLine = "AI news task started at " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $startLine | Tee-Object -FilePath $LogFile

    $process = Start-Process `
        -FilePath $Python `
        -ArgumentList @((Join-Path $ScriptDir "ai_news_daily.py"), "--output-dir", $OutputDir, "--report", $Report) `
        -WorkingDirectory $ScriptDir `
        -WindowStyle Hidden `
        -RedirectStandardOutput $StdoutFile `
        -RedirectStandardError $StderrFile `
        -Wait `
        -PassThru

    if (Test-Path -LiteralPath $StdoutFile) {
        Get-Content -LiteralPath $StdoutFile -Encoding UTF8 | Tee-Object -FilePath $LogFile -Append
    }
    if (Test-Path -LiteralPath $StderrFile) {
        Get-Content -LiteralPath $StderrFile -Encoding UTF8 | Tee-Object -FilePath $LogFile -Append
    }

    $exitCode = $process.ExitCode
    ("AI news task finished at " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + "; exit code: " + $exitCode) | Tee-Object -FilePath $LogFile -Append
    if ($exitCode -ne 0) {
        throw ("AI news report failed; exit code: " + $exitCode)
    }
} finally {
    Pop-Location
}
