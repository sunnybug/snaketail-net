# 功能说明：初始化开发环境（.run 目录、配置文件种子、还原 NuGet）

param()

$ErrorActionPreference = "Stop"
trap {
    Write-Host "命令行被中止: $_" -ForegroundColor Red
    Write-Host "$($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    exit 1
}

function Initialize-DevEnvironment {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $runDir = Join-Path $repoRoot ".run"
    $logDir = Join-Path $runDir "log"
    $configDir = Join-Path $runDir "config"
    foreach ($d in @($runDir, $logDir, $configDir)) {
        if (-not (Test-Path $d)) {
            New-Item -ItemType Directory -Path $d -Force | Out-Null
            Write-Host "已创建: $d" -ForegroundColor Green
        }
    }
    $srcCfg = Join-Path $repoRoot "src\app.config"
    $dstCfg = Join-Path $configDir "app.config"
    if ((Test-Path $srcCfg) -and -not (Test-Path $dstCfg)) {
        Copy-Item -LiteralPath $srcCfg -Destination $dstCfg -Force
        Write-Host "已复制默认配置: $dstCfg" -ForegroundColor Green
    }
    $sln = Join-Path $repoRoot "SnakeTail.sln"
    Write-Host "dotnet restore ..." -ForegroundColor Cyan
    dotnet restore $sln
    if ($LASTEXITCODE -ne 0) { throw "dotnet restore 失败" }
    Write-Host "开发环境初始化完成." -ForegroundColor Green
}

Initialize-DevEnvironment
