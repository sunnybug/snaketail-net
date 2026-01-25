# 编译并运行 SnakeTail
# 使用 UTF8 BOM CRLF 编码

$ErrorActionPreference = "Stop"

# 项目路径
$projectPath = "D:\xsw\code\snaketail-net"
$projectFile = "$projectPath\SnakeTail\SnakeTail.csproj"
$exePath = "$projectPath\SnakeTail\bin\Debug\net20\SnakeTail.exe"
$processName = "SnakeTail"

# MSBuild 路径
$msbuildPath = "D:\Apps\VisualStudio2022\MSBuild\Current\Bin\MSBuild.exe"

# 切换到项目目录
Set-Location $projectPath

# 强杀已运行的进程
Write-Host "正在检查并终止已运行的进程..." -ForegroundColor Yellow
$processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
if ($processes) {
    foreach ($proc in $processes) {
        Write-Host "正在终止进程: $($proc.Id) - $($proc.ProcessName)" -ForegroundColor Yellow
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Milliseconds 500
    Write-Host "进程已终止" -ForegroundColor Green
} else {
    Write-Host "没有发现运行中的进程" -ForegroundColor Green
}

Write-Host "正在编译项目..." -ForegroundColor Green

# 编译项目
& $msbuildPath $projectFile /p:Configuration=Debug /p:Platform=AnyCPU /t:Build

if ($LASTEXITCODE -ne 0) {
    Write-Host "编译失败！" -ForegroundColor Red
    exit 1
}

Write-Host "编译成功！" -ForegroundColor Green

# 检查 exe 是否存在
if (-not (Test-Path $exePath)) {
    Write-Host "找不到可执行文件: $exePath" -ForegroundColor Red
    exit 1
}

Write-Host "正在启动程序..." -ForegroundColor Green

# 运行程序
Start-Process $exePath

Write-Host "程序已启动" -ForegroundColor Green
