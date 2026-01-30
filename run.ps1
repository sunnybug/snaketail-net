# 编译并运行 SnakeTail
# 使用 UTF8 BOM CRLF 编码

$ErrorActionPreference = "Stop"

# 解析命令行参数
$debugMode = $false
foreach ($arg in $args) {
    if ($arg -eq "--debug" -or $arg -eq "-d") {
        $debugMode = $true
    }
}

# 项目路径
$projectPath = "D:\xsw\code\snaketail-net"
$projectFile = "$projectPath\SnakeTail\SnakeTail.csproj"
$exePath = "$projectPath\SnakeTail\bin\Debug\net8.0-windows\SnakeTail.exe"
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
if ($debugMode) {
    Write-Host "使用调试模式启动..." -ForegroundColor Cyan
    # 尝试使用 Visual Studio 调试器启动
    $devenvPath = "D:\Apps\VisualStudio2022\Common7\IDE\devenv.exe"
    if (Test-Path $devenvPath) {
        # 使用 devenv /DebugExe 启动调试器
        Write-Host "正在使用 Visual Studio 调试器启动..." -ForegroundColor Cyan
        & $devenvPath /DebugExe $exePath
    } else {
        # 如果找不到 devenv，尝试查找其他 Visual Studio 版本
        $vsPaths = @(
            "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe",
            "C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\devenv.exe",
            "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\devenv.exe",
            "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenv.exe",
            "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\Common7\IDE\devenv.exe",
            "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\devenv.exe"
        )
        $found = $false
        foreach ($path in $vsPaths) {
            if (Test-Path $path) {
                Write-Host "正在使用 Visual Studio 调试器启动: $path" -ForegroundColor Cyan
                & $path /DebugExe $exePath
                $found = $true
                break
            }
        }
        if (-not $found) {
            Write-Host "未找到 Visual Studio，使用普通模式启动（请手动附加调试器）" -ForegroundColor Yellow
            Start-Process $exePath
        }
    }
} else {
    Start-Process $exePath
}

Write-Host "程序已启动" -ForegroundColor Green
