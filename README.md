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

## 构建

- 常规 `dotnet build`（未显式指定 `OutDir`）时，主程序、测试项目、`LongMaidDisplayPlugin` 插件项目的编译中间文件与默认输出统一落到 `.temp/`。
- `.\script\build.ps1`：按 `Debug` 配置依次构建主程序与 `LongMaidDisplayPlugin` 插件，`SnakeTail.exe` 输出到 `.run\`，插件 DLL 输出到 `.run\config\plugins\龙女仆\`。
- `.\script\build.ps1 --release`：按 `Release` 配置依次构建主程序与插件，`SnakeTail.exe` 输出到 `.run\`，插件 DLL 输出到 `.run\config\plugins\龙女仆\`。
- 运行时目标固定为 `win-x64`，构建后会清理 `.run\**\runtimes\` 下非 `win` 与 `win-x64` 目录。
- 插件目录会在构建后净化，仅保留 `LongMaidDisplayPlugin.*` 与 `s_skill.json`，不会保留 `SnakeTail.*` 等主程序产物。

VS Code 推荐扩展（见 `.vscode/extensions.json`）：**C# Dev Kit**（调试、解决方案与测试）、**C#**（语言服务与代码分析）。

运行/调试：工作目录应为 `.run/`（`0run.ps1` 与 VS Code `launch.json` 已如此配置）。日志写入 `.run/log/`：`YYYY-MM-DD.log`（分级行格式）、`YYYY-MM-DD_crash.log`（未处理异常等）。

启动配置目录解析规则（启动恢复默认会话时）：优先使用工作目录下的 `config/`；若工作目录不存在 `config/`，则回退到 `SnakeTail.exe` 所在目录下的 `config/`。

## 轮询与空闲 CPU

- 文件日志窗口的默认轮询参数：`FileChangeCheckInterval=500ms`、`FileCheckInterval=10s`。
- 当旧会话配置缺失或写成 `<=0` 时，会自动回退到上述默认值，避免 WinForms `Timer` 落到 `100ms` 导致空闲期 CPU 偏高。
- 日志文件变化检测采用“事件驱动优先 + 低频轮询兜底”：`LogFileStream` 通过 `FileSystemWatcher` 接收目录事件，仅设置变更标记，在读取端以 `200ms` 防抖触发检查，减少空闲期无效 IO。
- 保留 `FileCheckInterval` 兜底校验，用于覆盖监听缓冲区溢出、网络盘事件延迟/丢失、重命名顺序异常等场景。
- 读取端在 EOF 时会额外按底层流长度执行一次轻量重同步：即使事件漏触发，也能在后续 tick 读取到追加的新行。
- 打开多个日志文件时会为每个标签页创建独立 Tail 配置副本，避免共享配置对象导致“文件监控串到最后一个文件”。
- 已修复快速过滤开启时的追尾读取基准：改为按“未过滤总行数”计算下一行，避免新增内容到达后界面长时间不刷新。

## 自动发布（CI）
修改 `src/SnakeTail.csproj` 中的版本号并推送到 `main`/`master` 后，Version Release 工作流会自动创建 tag。**要让 Publish 工作流被触发**，请在仓库 Settings → Secrets and variables → Actions 中新增 Secret：
- **名称**：`REPO_TOKEN`
- **值**：一个 Personal Access Token（需勾选 `repo` 权限）

未配置 `REPO_TOKEN` 时，tag 仍会由 `GITHUB_TOKEN` 创建并推送，但 GitHub 不会因该 push 触发 Publish；可改在 Actions 页手动运行 Publish 并输入 tag。

- 监控"大型"文本日志文件
- 监控 Windows 事件日志（无需管理员权限）
- 支持多种窗口模式（MDI、标签页、浮动窗口）
- 保存和加载整个窗口会话。可以在启动时通过命令行参数加载会话文件
  - “Open Session...” 支持在文件对话框中多选 `.xml` 会话并按选择顺序逐个加载
  - 多选打开文件/会话时，会弹出成功/失败汇总；失败原因会写入 `log/YYYY-MM-DD.log`
- 基于关键字匹配的句子高亮（支持正则表达式）
  - 关键字高亮：只高亮关键字文本本身，而不是整行背景
  - 行标识：匹配关键字的行在最左边显示一个色块标识
  - 支持快速高亮和配置的关键字高亮两种模式
- 使用键盘快捷键快速跳转到高亮的句子
- 切换书签并快速在书签间跳转
- 配置外部工具并绑定自定义快捷键（在高亮时触发执行）
- 跟踪循环日志，其中日志文件会定期截断/重命名
- 跟踪日志目录，显示最新的日志文件（支持通配符）
- 清空显示区域（快捷键 Ctrl+L）：记住当前读取的log文件位置，清空当前log的显示区域，从前一次读取的log文件位置继续读取；偏移量会优先采用读取器真实已读行号，避免大文件下清屏后重扫旧内容导致新日志显示延迟
- 在整个文本日志文件（或 Windows 事件日志）中搜索
  - 长时间全文搜索会定期处理界面消息，避免窗口看起来卡死
  - 未启用显示插件时优先直接搜索原始文本，减少大文件搜索的额外开销
  - 反向搜索改为单次顺序扫描后取最后命中，避免按递减行号反复从文件头重读
  - 搜索未命中改为搜索窗内联提示，不再弹模态框，避免提示不可见时把界面卡住
- 全局对话框（`MessageBox`）统一绑定 owner（主窗体/活动窗体/前台窗口兜底），减少弹窗被遮挡或假死观感
- 当检测到文件更改时，使用图标高亮窗口标签页
- 多文件同时监控时，非当前激活标签页若有内容变更，会在对应 tab 显示红点未读提示；切换到该页后自动清除
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

## 显示插件（TailForm）

- 作用域：仅文件日志窗口 `TailForm`，不扩展到 `EventLogForm`。
- 目录约定：`config/plugins/<插件名>/`。
- 发现规则：宿主启动/重载配置时扫描每个一级子目录中的 `*.dll`，查找首个实现 `ILogDisplayPlugin` 的类型。
- 菜单行为：右键菜单与主菜单同步出现“插件”子菜单，按勾选顺序决定插件执行顺序。
- 状态持久化：每个文件配置保存 `EnabledDisplayPlugins`，重启后打开同一配置会自动恢复启用顺序。
- 处理边界：插件只改“显示链路文本”和文本匹配行为，不改原始日志文件、原始行号和定位模型。

### 文本行为口径

- 基于处理后文本：
  - 列表显示
  - 快速过滤 / 反选过滤
  - 搜索（正向/反向）
  - 快速高亮与配置关键字高亮
  - 命中计数与外部工具触发
  - 右键菜单“复制（处理后文本）”
- 仍基于原始文本：
  - `Ctrl+C`（含行内双击选词后的快捷键复制）

### 插件触发规则（含多行块）

- 保持插件启用顺序语义：前面的插件先拿到输入并先决定是否 `Handled`。
- 新增可选块提取能力：实现 `ILogDisplayBlockPlugin` 的插件可在当前行命中后，向后收集完整文本块再进入 `CanProcess/TryProcess`。
- 当块插件排在前面且成功 `Handled` 时，后续单行插件不会再处理该次输入。
- 当块插件返回多行文本且与输入块行数一致时，宿主会按原始行号逐行回填显示，保持“原日志行 ↔ 处理后行”一一对应。
- 未实现块提取接口的插件保持原有单行行为，不受影响。

### 示例插件：龙女仆

- 插件目录：`config/plugins/龙女仆/`
- 配置文件：
  - `s_skill.json`：读取 `s_skill[*][0]=技能ID`、`s_skill[*][1]=技能名`
  - `s_battle_power.json`（可选）：读取 `s_battle_power[*][0]=key`、`s_battle_power[*][5]=名称`
- 处理规则：
  - 单行：命中 `skills: <数字>`、`passive_skill: <数字>` 或 `aura_skills: <数字>` 且存在映射时，显示为 `键名: <数字> <技能名>`；未知 ID（如 `passive_skill: 0`）保持原样。
  - 单行列表：命中 `skill: [<数字>,<数字>,...]` 或 `"skills": [<数字>,<数字>,...]` 时，按列表逐个扩展，已知 ID 显示为 `<数字> <技能名>`，未知 ID 保持原样（如 `"skills": [2001 名称A,3002,4001 名称B]`）。
  - 多行列表：命中 `"skills": [`（或 `skills: [`）起始并收集到 `]` 后，逐行将纯数字条目扩展为 `<数字> <技能名>`，保留原缩进与逗号。
  - 多行块：命中 `attr_data=effects {` 起始并收集连续 `effects` 结构块后，把 `key: <数字>` 扩展为 `key: <数字> <名称>`（如 `key: 1 声明`）。

