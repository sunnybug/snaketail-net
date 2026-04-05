# 功能说明：构建主程序与插件，默认 Debug，传 --release 时构建 Release

# 参数定义
param(
    [switch]$Release
)

# 兼容 --release 写法（PowerShell 7+ 会把 --release 解析为 $args 而非 switch）
if ($args -contains '--release' -or $args -contains 'release') { $Release = $true }

# 错误处理
$ErrorActionPreference = "Stop"
trap {
    Write-Host "命令行被中止: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "$($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    exit 1
}

# 路径与配置
$config = if ($Release) { "Release" } else { "Debug" }
$repoRoot = Split-Path -Parent $PSScriptRoot
$mainOutDir = Join-Path $repoRoot ".run"
$pluginOutDir = Join-Path $repoRoot ".run\config\plugins\龙女仆"
# 固定目标运行时为 Windows x64
$runtimeIdentifier = "win-x64"
# 兼容包内公共 Windows 资源目录
$runtimeKeepNames = @("win", "win-x64")

# 清理非 Windows x64 运行时目录
function Remove-NonWinX64Runtimes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetDir
    )

    $runtimesDir = Join-Path $TargetDir "runtimes"
    if (-not (Test-Path -LiteralPath $runtimesDir)) { return }
    Get-ChildItem -LiteralPath $runtimesDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        if (-not ($runtimeKeepNames -contains $_.Name)) {
            Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
        }
    }
}

# 对主目录与插件目录统一做最终清理
function Normalize-RuntimesForWinX64 {
    foreach ($dir in @($mainOutDir, $pluginOutDir)) {
        Remove-NonWinX64Runtimes -TargetDir $dir
    }
}

# 清理插件目录中的主程序及无关产物
function Keep-OnlyPluginArtifacts {
    $keepFileNames = @(
        "LongMaidDisplayPlugin.dll",
        "LongMaidDisplayPlugin.pdb",
        "LongMaidDisplayPlugin.deps.json",
        "s_skill.json"
    )
    if (-not (Test-Path -LiteralPath $pluginOutDir)) { return }
    Get-ChildItem -LiteralPath $pluginOutDir -Force -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.PSIsContainer) {
            Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
        } elseif (-not ($keepFileNames -contains $_.Name)) {
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
        }
    }
}

# 确保主程序与插件输出目录存在
foreach ($dir in @($mainOutDir, $pluginOutDir)) {
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}
$buildProjects = @(
    @{ Name = "主程序"; Path = (Join-Path $repoRoot "src\SnakeTail.csproj"); OutDir = $mainOutDir; ExtraArgs = @("/property:OutDir=$mainOutDir\", "/property:RuntimeIdentifier=$runtimeIdentifier", "/property:PlatformTarget=x64") },
    @{ Name = "龙女仆插件"; Path = (Join-Path $repoRoot "plugins\LongMaidDisplayPlugin\LongMaidDisplayPlugin.csproj"); OutDir = $pluginOutDir; ExtraArgs = @("/property:OutDir=$pluginOutDir\", "/property:RuntimeIdentifier=$runtimeIdentifier", "/property:PlatformTarget=x64") }
)

Write-Host "构建配置: $config" -ForegroundColor Cyan

# 依次构建主程序与插件
foreach ($item in $buildProjects) {
    if (-not (Test-Path -LiteralPath $item.Path)) {
        throw "构建失败：未找到项目文件（$($item.Name)）`n路径: $($item.Path)"
    }

    Write-Host "正在构建: $($item.Name)" -ForegroundColor Cyan
    # 产物输出到 .run 目录结构
    dotnet build $item.Path --configuration $config --nologo --verbosity quiet /property:GenerateFullPaths=true @($item.ExtraArgs)
    if ($LASTEXITCODE -ne 0) {
        throw "构建失败：$($item.Name)（退出码 $LASTEXITCODE）`n项目: $($item.Path)"
    }
}

# 清理无关平台运行时资产（最终态）
Normalize-RuntimesForWinX64
# 插件目录仅保留插件产物
Keep-OnlyPluginArtifacts

Write-Host "构建完成." -ForegroundColor Green
