$ErrorActionPreference = "Stop"

$TaskName = "AI News Daily Report"
$ScriptPath = Join-Path $PSScriptRoot "run_ai_news.ps1"
$WorkDir = $PSScriptRoot

if (-not (Test-Path $ScriptPath)) {
    throw "找不到运行脚本：$ScriptPath"
}

$Action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`"" `
    -WorkingDirectory $WorkDir

$Trigger = New-ScheduledTaskTrigger -Daily -At 9:00AM
$Principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME `
    -LogonType Interactive `
    -RunLevel Limited
$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Principal $Principal `
    -Settings $Settings `
    -Description "每天早上 9:00 自动生成 AI 相关新闻日报和大杂烩新闻日报到项目目录下的 output 文件夹" `
    -Force | Out-Null

Write-Host "已创建/更新计划任务：$TaskName"
Write-Host "每天执行时间：09:00"
Write-Host "输出目录：$PSScriptRoot\output"
