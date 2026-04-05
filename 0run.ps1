# 功能说明：编译、强杀目标程序、清除运行日志并启动 SnakeTail；支持 --test 运行单元测试

param(
    [switch]$Release,
    [switch]$Test
)

# 兼容 --release/--test 写法（PowerShell 7+ 会把 --xxx 解析为 $args 而非 switch）
$filterValue = $null
$i = 0
while ($i -lt $args.Count) {
    $a = $args[$i]
    if ($a -eq '--release' -or $a -eq 'release') { $Release = $true; $i++; continue }
    if ($a -eq '-Release') { $Release = $true; $i++; continue }
    if ($a -eq '--test') { $Test = $true; $i++; continue }
    if ($a -eq '-Test') { $Test = $true; $i++; continue }
    if ($a -eq '--filter') {
        if ($i + 1 -ge $args.Count) {
            Write-Host "错误: --filter 缺少参数值" -ForegroundColor Red
            Write-Host "用法: .\0run.ps1 [-Release] [--release] [--test] [--filter ""FullyQualifiedName~TestName""]" -ForegroundColor Yellow
            exit 1
        }
        $filterValue = $args[$i + 1]
        $i += 2
        continue
    }
    Write-Host "错误: 不支持的参数: $a" -ForegroundColor Red
    Write-Host "支持的参数: -Release, --release, --test, -Test, --filter ""表达式""" -ForegroundColor Yellow
    exit 1
}

$ErrorActionPreference = "Stop"
trap {
    Write-Host "命令行被中止: $_" -ForegroundColor Red
    Write-Host "$($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    exit 1
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$config = if ($Release) { "Release" } else { "Debug" }
$runDir = Join-Path $scriptRoot ".run"
$runLogDir = Join-Path $runDir "log"
$runConfigDir = Join-Path $runDir "config"

function Get-ProblemLogPaths {
    $paths = @()
    if (Test-Path $runLogDir) {
        Get-ChildItem -Path $runLogDir -File -ErrorAction SilentlyContinue | ForEach-Object {
            $fullName = $_.FullName
            $name = $_.Name
            $hasContent = $_.Length -gt 0
            if (-not $hasContent) { return }
            if ($name -match 'crash|error') { $paths += $fullName; return }
            if ($_.Extension -eq '.log') {
                $content = Get-Content -Path $fullName -Raw -ErrorAction SilentlyContinue
                if ($content -and $content -match 'error') { $paths += $fullName }
            }
        }
    }
    return $paths
}

# 1. 强杀目标程序
Write-Host "步骤 1/5: 强杀目标程序 ..." -ForegroundColor Cyan
Get-Process -Name "SnakeTail" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 300

# 2. 编译
Write-Host "步骤 2/5: 编译 ($config) ..." -ForegroundColor Cyan
$buildScript = Join-Path $scriptRoot "script\build.ps1"
if ($Release) { & $buildScript -Release } else { & $buildScript }
if ($LASTEXITCODE -ne 0) { throw "编译失败" }

# 3. 清除信息：工作目录 .run，清空 log；确保 config 并从 src 种子复制 app.config
Write-Host "步骤 3/5: 清除信息 ..." -ForegroundColor Cyan
if (-not (Test-Path $runDir)) { New-Item -ItemType Directory -Path $runDir -Force | Out-Null }
if (-not (Test-Path $runLogDir)) { New-Item -ItemType Directory -Path $runLogDir -Force | Out-Null }
Get-ChildItem -Path $runLogDir -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse
if (-not (Test-Path $runConfigDir)) { New-Item -ItemType Directory -Path $runConfigDir -Force | Out-Null }
$srcAppConfig = Join-Path $scriptRoot "src\app.config"
$dstAppConfig = Join-Path $runConfigDir "app.config"
if ((Test-Path $srcAppConfig) -and -not (Test-Path $dstAppConfig)) {
    Copy-Item -LiteralPath $srcAppConfig -Destination $dstAppConfig -Force
}

# 4. 运行
Write-Host "步骤 4/5: 运行 ..." -ForegroundColor Cyan
# 若项目为插件则安装到目标程序中（当前为主程序，不执行）

if ($Test) {
    $testsProj = $null
    $testsDir = Join-Path $scriptRoot "tests"
    if (Test-Path $testsDir) {
        $testsProj = Get-ChildItem -Path $testsDir -Filter "*.csproj" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if (-not $testsProj) {
        Write-Host "错误: 暂无单元测试项目（未找到 tests 目录下的 .csproj）" -ForegroundColor Red
        exit 1
    }
    $testArgs = @("test", $testsProj.FullName, "--configuration", $config)
    if ($filterValue) { $testArgs += "--filter"; $testArgs += $filterValue }
    & dotnet $testArgs
    $testExitCode = $LASTEXITCODE
    Write-Host "步骤 5/5: 检查结果 ..." -ForegroundColor Cyan
    $problemLogs = Get-ProblemLogPaths
    if ($problemLogs.Count -gt 0) {
        foreach ($p in $problemLogs) { Write-Host $p }
    }
    if ($testExitCode -ne 0) { exit $testExitCode }
    exit 0
}

$exe = Join-Path $scriptRoot ".run\SnakeTail.exe"
if (-not (Test-Path $exe)) {
    throw "未找到可执行文件: $exe"
}
Start-Process -FilePath $exe -WorkingDirectory $runDir -Wait

Write-Host "步骤 5/5: 检查结果 ..." -ForegroundColor Cyan
$problemLogs = Get-ProblemLogPaths
if ($problemLogs.Count -gt 0) {
    foreach ($p in $problemLogs) { Write-Host $p }
}
