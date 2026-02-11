# 功能说明：构建 SnakeTail 项目，默认 Debug，传 --release 时构建 Release

param(
    [switch]$Release
)
# 兼容 --release 写法（PowerShell 7+ 会把 --release 解析为 $args 而非 switch）
if ($args -contains '--release' -or $args -contains 'release') { $Release = $true }

$ErrorActionPreference = "Stop"
trap {
    Write-Host "命令行被中止: $_" -ForegroundColor Red
    Write-Host "$($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    exit 1
}

$config = if ($Release) { "Release" } else { "Debug" }
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$proj = Join-Path $scriptRoot "SnakeTail\SnakeTail.csproj"

Write-Host "构建配置: $config" -ForegroundColor Cyan
dotnet build $proj --configuration $config /property:GenerateFullPaths=true "/consoleloggerparameters:NoSummary;ForceNoAlign"
if ($LASTEXITCODE -ne 0) { throw "构建失败" }
Write-Host "构建完成." -ForegroundColor Green
