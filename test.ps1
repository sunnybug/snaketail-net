# 功能说明：编译、强杀目标程序、清除运行日志并启动 SnakeTail

param(
    [switch]$Release
)
# 兼容 --release 写法（PowerShell 7+ 会把 --release 解析为 $args 而非 switch）
if ($args -contains '--release' -or $args -contains 'release') { $Release = $true }

$ErrorActionPreference = "Stop"
trap {
    Write-Host "命令行被中止: $_" -ForegroundColor Red
    Write-Host "$($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    Read-Host "按 Enter 键关闭窗口"
    break
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$config = if ($Release) { "Release" } else { "Debug" }

# 1. 编译
Write-Host "步骤 1/4: 编译 ($config) ..." -ForegroundColor Cyan
$buildScript = Join-Path $scriptRoot "build.ps1"
if ($Release) { & $buildScript -Release } else { & $buildScript }
if ($LASTEXITCODE -ne 0) { throw "编译失败" }

# 2. 强杀目标程序
Write-Host "步骤 2/4: 强杀目标程序 ..." -ForegroundColor Cyan
Get-Process -Name "SnakeTail" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 300

# 3. 清除运行日志
Write-Host "步骤 3/4: 清除运行日志 ..." -ForegroundColor Cyan
$tempLogs = Join-Path $env:TEMP "SnakeTail_*.txt"
Get-ChildItem -Path $tempLogs -ErrorAction SilentlyContinue | Remove-Item -Force
$runLogDir = Join-Path $scriptRoot ".run\log"
if (Test-Path $runLogDir) {
    Get-ChildItem -Path $runLogDir -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse
}

# 4. 启动目标程序
Write-Host "步骤 4/4: 启动目标程序 ..." -ForegroundColor Cyan
$exeDir = Join-Path $scriptRoot "SnakeTail\bin\$config\net8.0-windows"
$exe = Join-Path $exeDir "SnakeTail.exe"
if (-not (Test-Path $exe)) {
    throw "未找到可执行文件: $exe"
}
Start-Process -FilePath $exe -WorkingDirectory $exeDir
Write-Host "已启动 SnakeTail." -ForegroundColor Green
