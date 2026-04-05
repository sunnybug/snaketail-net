# 功能说明：构建 Release 并将可发布文件复制到仓库根目录 .dist 下（带版本号子目录）

param()

$ErrorActionPreference = "Stop"
trap {
    Write-Host "命令行被中止: $_" -ForegroundColor Red
    Write-Host "$($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    exit 1
}

function Publish-ReleaseToDist {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $proj = Join-Path $repoRoot "src\SnakeTail.csproj"
    $raw = Get-Content -LiteralPath $proj -Raw -Encoding UTF8
    if ($raw -notmatch '<Version>([^<]+)</Version>') {
        throw "无法从 csproj 解析 <Version>"
    }
    $version = $Matches[1].Trim()
    $buildScript = Join-Path $repoRoot "script\build.ps1"
    & $buildScript -Release
    if ($LASTEXITCODE -ne 0) { throw "Release 构建失败" }

    $outDir = Join-Path $repoRoot ".temp\bin\Release\net8.0-windows"
    if (-not (Test-Path (Join-Path $outDir "SnakeTail.exe"))) {
        throw "未找到构建输出: $outDir\SnakeTail.exe"
    }
    $distRoot = Join-Path $repoRoot ".dist"
    $distDir = Join-Path $distRoot "SnakeTail-$version"
    if (Test-Path $distDir) {
        Remove-Item -LiteralPath $distDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $distDir -Force | Out-Null
    Copy-Item -Path (Join-Path $outDir '*') -Destination $distDir -Recurse -Force
    Write-Host "已发布到: $distDir" -ForegroundColor Green
}

Publish-ReleaseToDist
