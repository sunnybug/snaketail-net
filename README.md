# snaketail-net
用于监控文本日志文件和 Windows 事件日志的 Tail 工具

## 目录结构（约定）

| 路径 | 说明 |
|------|------|
| `src/` | 主程序源码与 `SnakeTail.csproj` |
| `tests/` | 单元测试 |
| `script/` | `build.ps1`、`publish.ps1`、`init_dev.ps1` |
| `.temp/` | 编译中间文件与输出（由 csproj 重定向，勿提交） |
| `.run/` | 本地运行工作目录（`log/`、`config/`，勿提交） |
| `.dist/` | `publish.ps1` 生成的本地发布目录（勿提交） |
| `0run.ps1` | 一键编译、清理 `.run/log`、从 `src\app.config` 种子复制到 `.run\config`（若缺失）、启动或跑测试 |

开发时请先执行 `.\script\init_dev.ps1` 初始化 `.run` 并 `dotnet restore`。

VS Code 推荐扩展（见 `.vscode/extensions.json`）：**C# Dev Kit**（调试、解决方案与测试）、**C#**（语言服务与代码分析）。

运行/调试：工作目录应为 `.run/`（`0run.ps1` 与 VS Code `launch.json` 已如此配置）。日志写入 `.run/log/`：`YYYY-MM-DD.log`（分级行格式）、`YYYY-MM-DD_crash.log`（未处理异常等）。

## 自动发布（CI）
修改 `src/SnakeTail.csproj` 中的版本号并推送到 `main`/`master` 后，Version Release 工作流会自动创建 tag。**要让 Publish 工作流被触发**，请在仓库 Settings → Secrets and variables → Actions 中新增 Secret：
- **名称**：`REPO_TOKEN`
- **值**：一个 Personal Access Token（需勾选 `repo` 权限）

未配置 `REPO_TOKEN` 时，tag 仍会由 `GITHUB_TOKEN` 创建并推送，但 GitHub 不会因该 push 触发 Publish；可改在 Actions 页手动运行 Publish 并输入 tag。

- 监控"大型"文本日志文件
- 监控 Windows 事件日志（无需管理员权限）
- 支持多种窗口模式（MDI、标签页、浮动窗口）
- 保存和加载整个窗口会话。可以在启动时通过命令行参数加载会话文件
- 基于关键字匹配的句子高亮（支持正则表达式）
  - 关键字高亮：只高亮关键字文本本身，而不是整行背景
  - 行标识：匹配关键字的行在最左边显示一个色块标识
  - 支持快速高亮和配置的关键字高亮两种模式
- 使用键盘快捷键快速跳转到高亮的句子
- 切换书签并快速在书签间跳转
- 配置外部工具并绑定自定义快捷键（在高亮时触发执行）
- 跟踪循环日志，其中日志文件会定期截断/重命名
- 跟踪日志目录，显示最新的日志文件（支持通配符）
- 清空显示区域（快捷键 Ctrl+L）：记住当前读取的log文件位置，清空当前log的显示区域，从前一次读取的log文件位置继续读取
- 在整个文本日志文件（或 Windows 事件日志）中搜索
- 当检测到文件更改时，使用图标高亮窗口标签页
- 通过简单的拖放操作从 Windows 资源管理器跟踪新日志文件
- 使用正则表达式过滤 Windows 事件日志
- 在窗口标题栏显示简单的进程统计信息（内存 + CPU 使用率 + 事务/秒）
- 直接停止和启动服务
- 更改跟踪窗口背景颜色
- 更改跟踪窗口文本颜色
- 更改跟踪窗口图标
- 最小化到系统托盘
- 低内存占用，与日志文件大小无关
- 即使每秒超过 100 行也能保持低 CPU 使用率
- 在远程桌面上运行良好
- 支持 Windows 2000、XP、2003、Vista、Win2k8、Win7
- 需要 .NET 2.0
- GNU GPL 许可证 v3
