# 通配符文件监控与规则对话框

## TL;DR
> **Summary**: 在现有 WinForms 架构上新增“目录 + 通配符”监控规则能力：用户在独立规则对话框里配置单层目录与 `*`/`?` 模式，系统自动发现并接管匹配到的文件，复用现有 `TailForm` 打开/激活标签页。
> **Deliverables**:
> - 会话级 `MonitorRuleConfig` 数据模型与持久化
> - 规则校验/预览逻辑
> - 基于 `FileSystemWatcher + Directory.GetFiles` 的规则运行时管理器
> - `MainForm` 菜单入口与独立“Monitor Rules”对话框
> - README / CLAUDE 使用说明与回归验证
> **Effort**: Large
> **Parallel**: YES - 2 waves
> **Critical Path**: 1 → 3 → 4 → 5 → 6/7 → 8 → 9

## Context
### Original Request
- 新增功能：对指定目录下的匹配通配符的文件监控，例如 `d:\xxx\*err*.log`
- 需要同时给出方案与界面

### Interview Summary
- 主应用已确认是 **.NET 8 WinForms**，入口为 `Program -> MainForm`
- 范围限定为：**仅当前目录**，**不递归子目录**
- 匹配语法限定为：**仅 `*` 与 `?`**，不支持正则、不支持排除规则
- 规则命中新文件后：**自动纳入监控**，不要求用户再次确认
- UI 方向已定：`MainForm` 提供入口，使用**独立规则管理对话框**
- 测试策略已定：**TDD**

### Metis Review (gaps addressed)
- 补齐并固定以下默认决策，避免执行时再做判断：
  - 通配规则只匹配“文件名部分”，目录单独存储；UI 里展示组合预览路径
  - Windows 下按**不区分大小写**处理路径与文件名匹配
  - 规则启用时先做一次**现有文件发现**，随后再接 watcher 事件
  - 文件被删除/重命名后**不自动关闭**已打开的 `TailForm`，沿用现有 Tail 行为处理缺失文件
  - 若匹配文件已在当前窗口打开，则**激活现有 Tab，不创建重复 Tab**
- guardrails：
  - **不得改坏**现有 `TailFileConfig.FileCheckPattern` 的“最新匹配文件”旧语义
  - **不引入** DI 容器、递归扫描、正则匹配、逐条命中弹确认框

## Work Objectives
### Core Objective
在不破坏现有单文件 Tail 与旧通配路径语义的前提下，为 WinForms 主程序增加一套“监控规则”能力：用户可持续监控指定目录下符合模式的多个文件，并自动以现有 Tail 视图纳入会话。

### Deliverables
- `MonitorRuleConfig` 会话模型
- `MonitorRuleValidation` 校验/预览逻辑
- `MonitorRuleManager` 运行时发现与 watcher 协调逻辑
- `MonitorRuleEditForm`（新增/编辑单条规则）
- `MonitorRulesForm`（规则管理与命中状态总览）
- `MainForm` 菜单入口、启动/关闭生命周期接入
- 文档更新：`README.md`、`CLAUDE.md`

### Definition of Done (verifiable conditions with commands)
- 执行 `pwsh -File ".\script\build.ps1" --release` 前，确保运行中的 `SnakeTail.exe` 已退出，避免 `.run\SnakeTail.exe` 被锁定导致 `MSB3027/MSB3021`
- `dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRule"`
- `dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~LogFileStream"`
- `pwsh -File ".\script\build.ps1" --release`
- 手动/代理 QA 通过以下关键场景：
  - 新建规则后能立即发现当前目录已存在的匹配文件
  - 新文件落入目录后会自动打开或激活对应标签页
  - 同一文件不会因重复事件被打开多个标签页
  - 保存会话/重启后规则仍能恢复

### Must Have
- 新规则字段：`Name`、`DirectoryPath`、`FilePattern`、`Enabled`
- 规则启用后立即扫描当前目录一次，再接入实时 watcher
- 复用 `MainForm.OpenFileOrActivateTab` / `BringToFrontAndOpenOrActivateTab` 的既有打开/激活逻辑
- UI 中提供：新增、编辑、删除、启用/停用、查看当前匹配文件状态
- 编辑对话框提供组合预览：`<DirectoryPath>\<FilePattern>`
- 规则保存到 `TailConfig`，跟随 XML 会话与默认会话一并持久化

### Must NOT Have (guardrails, AI slop patterns, scope boundaries)
- 不支持递归子目录
- 不支持正则、排除模式、多目录规则
- 不引入新的依赖注入框架或重构全局架构
- 不改变旧的 `TailConfigForm` / `TailFileConfig.FileCheckPattern` 现有行为口径
- 不对每个新命中文件弹确认框
- 不在文件删除时自动关闭用户已有的 Tail 标签页

## Verification Strategy
> ZERO HUMAN INTERVENTION - all verification is agent-executed.
- Test decision: **TDD** + xUnit（已有 `tests/SnakeTail.Tests/SnakeTail.Tests.csproj`）
- QA policy: 每个任务都附带 happy / failure 场景；UI 场景通过可测试逻辑 + 最终代理 QA 收口
- Evidence: `.sisyphus/evidence/task-{N}-{slug}.{ext}`

## Execution Strategy
### Parallel Execution Waves
> Target: 5-8 tasks per wave. <3 per wave (except final) = under-splitting.
> Extract shared dependencies as Wave-1 tasks for max parallelism.

Wave 1: 数据契约与运行时基础
- T1 MonitorRuleConfig 数据模型
- T2 规则校验与组合预览
- T3 目录扫描 / watcher 封装
- T4 规则管理器与重复命中去重
- T5 MainForm 生命周期与会话接入

Wave 2: UI 与收尾
- T6 MonitorRuleEditForm 单条规则编辑对话框
- T7 MonitorRulesForm 规则管理总对话框
- T8 MainForm 菜单接入 + 对话框状态持久化
- T9 README / CLAUDE / 最终回归脚本与说明

### Dependency Matrix (full, all tasks)
| Task | Depends On | Needed By |
|---|---|---|
| T1 | - | T2, T3, T4, T5, T6, T7, T9 |
| T2 | T1 | T3, T4, T6, T7 |
| T3 | T1, T2 | T4, T5 |
| T4 | T1, T2, T3 | T5, T7 |
| T5 | T1, T3, T4 | T8, T9 |
| T6 | T1, T2 | T7 |
| T7 | T4, T6 | T8 |
| T8 | T5, T7 | T9 |
| T9 | T5, T8 | Final Verification |

### Agent Dispatch Summary (wave → task count → categories)
- Wave 1 → 5 tasks → `quick` ×2, `unspecified-high` ×3
- Wave 2 → 4 tasks → `visual-engineering` ×2, `unspecified-high` ×1, `writing` ×1
- Final Verification → 4 tasks → `oracle` / `unspecified-high` / `deep`

## TODOs
> Implementation + Test = ONE task. Never separate.
> EVERY task MUST have: Agent Profile + Parallelization + QA Scenarios.

- [x] 1. 扩展会话模型以持久化监控规则

  **What to do**:
  - 在 `src/TailConfig.cs` 中新增 `MonitorRuleConfig`，字段固定为：
    - `string Name`
    - `string DirectoryPath`
    - `string FilePattern`
    - `bool Enabled`
  - 在 `TailConfig` 上新增 `List<MonitorRuleConfig> MonitorRules`，构造时初始化为空列表，确保旧会话反序列化后不会得到 `null`
  - `Name` 允许为空；显示时若为空，统一由 UI 逻辑回退为 `Path.GetFileName(DirectoryPath) + " | " + FilePattern`
  - **不要**把每条规则扩展成 Tail 外观模板；首版统一复用当前默认 Tail 配置
  - 新增 xUnit 回归测试覆盖：默认值、安全反序列化、旧 XML 不含 `MonitorRules` 字段时仍能加载

  **Must NOT do**:
  - 不修改现有 `TailFileConfig` 字段含义
  - 不把规则存到独立 JSON / SQLite 表；首版只跟随 `TailConfig` 会话存储

  **Recommended Agent Profile**:
  - Category: `quick` - Reason: 单文件数据契约扩展 + 对应测试，判断简单
  - Skills: [`superpowers:test-driven-development`] - 先写反序列化失败测试，再改 DTO
  - Omitted: [`my-ui`] - 本任务无界面实现

  **Parallelization**: Can Parallel: YES | Wave 1 | Blocks: [2, 3, 4, 5, 6, 7, 9] | Blocked By: []

  **References**:
  - Pattern: `src/TailConfig.cs:25-35` - 会话级配置容器 `TailConfig`
  - Pattern: `src/TailConfig.cs:171-219` - 现有 `TailFileConfig` 序列化字段风格
  - Pattern: `tests/SnakeTail.Tests/TailFileConfigDefaultsTests.cs:8-31` - 默认值/旧 XML 回归测试写法
  - API/Type: `src/MainForm.cs:916-930` - `XmlSerializer(typeof(TailConfig))` 的实际会话保存路径
  - API/Type: `src/Storage/SnakeTailStorage.cs:204-230` - 默认会话通过同一 `TailConfig` 序列化/反序列化

  **Acceptance Criteria**:
  - [ ] `TailConfig` 拥有非空 `MonitorRules` 集合，旧会话 XML 反序列化后不会抛错也不会得到 `null`
  - [ ] 新增测试文件 `tests/SnakeTail.Tests/MonitorRuleConfigTests.cs`
  - [ ] `dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRuleConfigTests"` 通过

  **QA Scenarios**:
  ```
  Scenario: 新会话模型能安全序列化/反序列化
    Tool: Bash
    Steps: 运行 dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRuleConfigTests"
    Expected: 所有测试 PASS，覆盖默认值与 XML round-trip
    Evidence: .sisyphus/evidence/task-1-monitor-rule-config.txt

  Scenario: 旧 XML 缺失 MonitorRules 字段时仍兼容
    Tool: Bash
    Steps: 在同一测试集中运行缺失字段兼容用例
    Expected: 反序列化成功，MonitorRules 非 null 且 Count == 0
    Evidence: .sisyphus/evidence/task-1-monitor-rule-config-error.txt
  ```

  **Commit**: YES | Message: `feat(config): add monitor rule session model` | Files: [`src/TailConfig.cs`, `tests/SnakeTail.Tests/MonitorRuleConfigTests.cs`]

- [x] 2. 固定规则校验、规范化与组合预览语义

  **What to do**:
  - 新建 `src/MonitorRuleValidation.cs`，提供纯逻辑静态/内部方法：
    - `NormalizeDirectoryPath(string)`：去首尾空白，转绝对路径，去掉末尾分隔符
    - `ValidateDirectory(string)`：目录不能为空；目录不存在时返回明确错误文本
    - `ValidateFilePattern(string)`：不能为空；不能包含目录分隔符；仅允许普通字符、`*`、`?`
    - `BuildPreviewPath(string directoryPath, string filePattern)`：输出 `<dir>\<pattern>`
    - `BuildDisplayName(MonitorRuleConfig)`：空名称时回退默认显示名
  - 明确首版语义：`FilePattern` 只对应文件名，不允许输入 `d:\a\b\*.log` 这种完整路径到 pattern 框；完整路径只在预览文本中展示
  - 新增 `tests/SnakeTail.Tests/MonitorRuleValidationTests.cs` 覆盖：合法模式、空目录、目录不存在、pattern 带分隔符、大小写/空白规范化

  **Must NOT do**:
  - 不支持 `**`、排除模式、正则
  - 不把目录不存在 silently 修正为最近存在的父目录

  **Recommended Agent Profile**:
  - Category: `quick` - Reason: 独立纯逻辑 helper，适合 TDD
  - Skills: [`superpowers:test-driven-development`] - 先锁定错误消息和 preview 语义
  - Omitted: [`my-ui`] - 本任务不创建窗体

  **Parallelization**: Can Parallel: YES | Wave 1 | Blocks: [3, 4, 6, 7] | Blocked By: [1]

  **References**:
  - Pattern: `src/TailConfigForm.cs:61-188` - 现有配置对话框的“载入/保存前做值转换”风格
  - Pattern: `src/MainForm.cs:973-985` - 错误消息要可精确追踪，沿用明确文本拼接风格
  - Test: `tests/SnakeTail.Tests/LogFileStreamWatcherMatchTests.cs:8-45` - Windows 下大小写不敏感的文件匹配口径
  - Test: `tests/SnakeTail.Tests/TailFileConfigDefaultsTests.cs:8-31` - xUnit 纯逻辑测试风格

  **Acceptance Criteria**:
  - [ ] `MonitorRuleValidation` 对非法目录和非法 pattern 返回稳定、可断言的错误文本
  - [ ] 预览路径统一输出 `DirectoryPath + "\\" + FilePattern`
  - [ ] `dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRuleValidationTests"` 通过

  **QA Scenarios**:
  ```
  Scenario: 合法目录+模式能生成组合预览
    Tool: Bash
    Steps: 运行 dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRuleValidationTests"
    Expected: preview、display-name、normalize 测试全部 PASS
    Evidence: .sisyphus/evidence/task-2-monitor-rule-validation.txt

  Scenario: 非法 pattern 被拦截且给出精确原因
    Tool: Bash
    Steps: 在同一测试集中执行 pattern 带路径分隔符、空串、非法目录场景
    Expected: 测试断言错误文本稳定，不接受无提示失败
    Evidence: .sisyphus/evidence/task-2-monitor-rule-validation-error.txt
  ```

  **Commit**: YES | Message: `feat(monitor): add rule validation helpers` | Files: [`src/MonitorRuleValidation.cs`, `tests/SnakeTail.Tests/MonitorRuleValidationTests.cs`]

- [x] 3. 新增目录扫描与 watcher 封装层

  **What to do**:
  - 新建 `src/MonitorDirectoryWatcher.cs`，封装“单条规则 -> 单个目录 watcher”职责
  - 固定实现策略：
    - 启用规则时先执行 `Directory.GetFiles(directory, pattern, TopDirectoryOnly)` 做基线发现
    - 再创建 `FileSystemWatcher`
    - `Path = DirectoryPath`
    - `Filter = FilePattern`
    - `IncludeSubdirectories = false`
    - 监听 `Created`、`Renamed`、`Changed`
  - 输出统一事件：`MatchFound(string absolutePath)`、`RuleError(string message)`
  - `Changed` 事件只用于“文件先创建后写入”的兜底，不得导致重复洪泛；同一路径在短时间窗口内要做去抖（建议 300ms 内按路径去重）
  - 目录不存在时不抛 UI 级异常，改为 `RuleError` 并标记规则不可运行
  - 新增 `tests/SnakeTail.Tests/MonitorDirectoryWatcherTests.cs`：覆盖初始发现、创建后命中、目录缺失错误

  **Must NOT do**:
  - 不复用 `LogFileStream.FindFileUsingPattern` 的“只找最新文件”逻辑
  - 不让 watcher 直接 new `TailForm`

  **Recommended Agent Profile**:
  - Category: `unspecified-high` - Reason: 文件系统事件有时序与去重细节
  - Skills: [`superpowers:test-driven-development`] - 先锁定初始发现与去抖测试
  - Omitted: [`my-ui`] - 纯运行时逻辑

  **Parallelization**: Can Parallel: YES | Wave 1 | Blocks: [4, 5, 7] | Blocked By: [1, 2]

  **References**:
  - Pattern: `src/LogFileStream.cs:46-58` - 现有 `FileSystemWatcher` 构造与初始化入口
  - Pattern: `src/LogFileStream.cs:70-149` - 文件变更后的 reload/check 处理思路
  - Pattern: `src/LogFileStream.cs:220-237` - 旧通配模式只找最新文件；本任务必须绕开该旧语义
  - Pattern: `src/LogFileStream.cs:557-580` - 现有事件驱动 + 去抖思路可参考
  - Test: `tests/SnakeTail.Tests/LogFileStreamAppendDetectionTests.cs:10-67` - 基于临时目录/临时文件的文件测试写法
  - Test: `tests/SnakeTail.Tests/LogFileStreamWatcherMatchTests.cs:34-44` - 通配 watcher 语义参考

  **Acceptance Criteria**:
  - [ ] 启用规则时能先返回当前目录已有匹配文件
  - [ ] 新文件创建或重命名进入目录后，能在去抖后仅上报一次 `MatchFound`
  - [ ] 缺失目录不会崩溃，错误通过 `RuleError` 暴露
  - [ ] `dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorDirectoryWatcherTests"` 通过

  **QA Scenarios**:
  ```
  Scenario: 启用规则时立即发现现有匹配文件
    Tool: Bash
    Steps: 运行 dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorDirectoryWatcherTests"
    Expected: 基线发现与新建文件命中测试 PASS
    Evidence: .sisyphus/evidence/task-3-monitor-directory-watcher.txt

  Scenario: 目录缺失时返回错误而非抛出未处理异常
    Tool: Bash
    Steps: 在同一测试集中执行不存在目录场景
    Expected: 断言收到 RuleError；测试进程不崩溃
    Evidence: .sisyphus/evidence/task-3-monitor-directory-watcher-error.txt
  ```

  **Commit**: YES | Message: `feat(monitor): add directory watcher` | Files: [`src/MonitorDirectoryWatcher.cs`, `tests/SnakeTail.Tests/MonitorDirectoryWatcherTests.cs`]

- [ ] 4. 实现规则管理器与重复命中去重注册表

  **What to do**:
  - 新建 `src/MonitorRuleManager.cs`，集中管理所有 `MonitorDirectoryWatcher`
  - 固定职责：
    - `ApplyRules(IReadOnlyList<MonitorRuleConfig>)`
    - `StartAll()` / `StopAll()` / `Dispose()`
    - 维护 `RuleState`（Running / Disabled / Error）与 `MatchedFiles` 快照
    - 对外发出 `RuleMatchDiscovered(ruleId/name, absolutePath)` 与 `RulesChanged` 事件
  - 内部维护**大小写不敏感**的绝对路径集合，用于抑制同一文件的重复事件
  - 同一文件若已由另一条规则管理，也只发一次“需要打开/激活”的事件；UI 层只显示它被多个规则命中即可，不重复开标签
  - 规则禁用时停止对应 watcher，但**不关闭**已打开的 `TailForm`
  - 新增 `tests/SnakeTail.Tests/MonitorRuleManagerTests.cs` 覆盖：重复命中去重、跨规则共享文件、禁用规则后不再继续发事件

  **Must NOT do**:
  - 不把 `MainForm` 或 `TailForm` 直接塞进 manager
  - 不在 manager 内弹 `MessageBox`

  **Recommended Agent Profile**:
  - Category: `unspecified-high` - Reason: 需要协调多规则、多 watcher、多事件去重
  - Skills: [`superpowers:test-driven-development`] - 先锁定多规则重复命中测试
  - Omitted: [`my-ui`] - 仍是非 UI 核心逻辑

  **Parallelization**: Can Parallel: YES | Wave 1 | Blocks: [5, 7] | Blocked By: [1, 2, 3]

  **References**:
  - Pattern: `src/MainForm.cs:546-617` - 已有“打开或激活现有标签页”口径，manager 只负责发事件不直接操纵 UI
  - Pattern: `src/LogFileStream.cs:38-45` - 现有事件模型示例（FileReloadedEvent / FileChangedEvent）
  - Test: `tests/SnakeTail.Tests/LogFileStreamWatcherMatchTests.cs:8-45` - 大小写不敏感匹配回归口径
  - Test: `tests/SnakeTail.Tests/LogFileStreamAppendDetectionTests.cs:35-66` - 临时目录/等待工具辅助写法

  **Acceptance Criteria**:
  - [ ] 同一路径被多次事件命中时只上报一次待打开事件
  - [ ] 同一路径被多规则同时命中时不会要求 UI 打开多个标签页
  - [ ] 停用规则只停止后续发现，不会关闭已打开 tail 窗口
  - [ ] `dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRuleManagerTests"` 通过

  **QA Scenarios**:
  ```
  Scenario: 多个 watcher 对同一文件重复命中时只触发一次打开请求
    Tool: Bash
    Steps: 运行 dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRuleManagerTests"
    Expected: 去重与停用规则测试全部 PASS
    Evidence: .sisyphus/evidence/task-4-monitor-rule-manager.txt

  Scenario: 规则停用后不再继续发送新命中
    Tool: Bash
    Steps: 在同一测试集中执行停用后新建文件场景
    Expected: 不再收到新的 match-discovered 事件，已知文件状态仍保留
    Evidence: .sisyphus/evidence/task-4-monitor-rule-manager-error.txt
  ```

  **Commit**: YES | Message: `feat(monitor): add monitor rule manager` | Files: [`src/MonitorRuleManager.cs`, `tests/SnakeTail.Tests/MonitorRuleManagerTests.cs`]

- [ ] 5. 将规则管理器接入 MainForm 生命周期与会话恢复

  **What to do**:
  - 修改 `src/MainForm.cs`：
    - 新增字段保存 `MonitorRuleManager`
    - 在 `LoadSessionWithConfig` 完成 `TailFiles` 加载后，应用并启动 `tailConfig.MonitorRules`
    - 在 `BuildCurrentTailConfig` / `SaveSession` / `SaveSessionToDb` 中把当前规则写回 `TailConfig.MonitorRules`
    - 在 `MainForm_FormClosing` 中先 `StopAll/Dispose`，再走现有会话保存
  - 规则命中新文件时，统一调用 `BringToFrontAndOpenOrActivateTab(new[]{path})`；**不要**自己复制一套打开逻辑
  - 若规则命中文件已经存在于当前 `TailForm` 标签中，必须沿用 `OpenFileOrActivateTab` 的去重激活语义
  - 当恢复默认会话或 XML 会话时，规则应跟着会话自动恢复
  - 新增 `tests/SnakeTail.Tests/MonitorRuleSessionPersistenceTests.cs`，至少覆盖 `TailConfig` round-trip 后规则仍在；若 MainForm 难以直测，则抽出可测 helper 后再单测 helper，不要硬写脆弱窗体测试

  **Must NOT do**:
  - 不直接在 `Program.cs` 挂规则逻辑
  - 不在 `LoadSessionWithConfig` 之前抢先启动规则，避免在 tail 窗口恢复前乱序开窗

  **Recommended Agent Profile**:
  - Category: `unspecified-high` - Reason: 涉及启动顺序、会话恢复和现有开窗逻辑复用
  - Skills: [`superpowers:test-driven-development`] - 先加持久化回归测试，再接 MainForm 生命周期
  - Omitted: [`my-ui`] - 主要是控制流接线

  **Parallelization**: Can Parallel: YES | Wave 1 | Blocks: [8, 9] | Blocked By: [1, 3, 4]

  **References**:
  - Pattern: `src/Program.cs:54-64` - 入口层只 new `MainForm`，不应扩散功能逻辑到 Program
  - Pattern: `src/MainForm.cs:546-617` - `OpenFileOrActivateTab` / `BringToFrontAndOpenOrActivateTab` 现成复用点
  - Pattern: `src/MainForm.cs:839-944` - XML 会话保存入口
  - Pattern: `src/MainForm.cs:1043-1090` - 默认会话保存构造逻辑
  - Pattern: `src/MainForm.cs:1092-1175` - 会话恢复顺序
  - Pattern: `src/MainForm.cs:1491-1558` - 关闭时保存/释放生命周期
  - API/Type: `src/Storage/SnakeTailStorage.cs:183-230` - 默认会话持久化路径

  **Acceptance Criteria**:
  - [ ] 保存 XML 会话与默认会话时，`MonitorRules` 均被持久化
  - [ ] 恢复会话后，启用中的规则自动启动并接管匹配文件
  - [ ] 同一路径命中时仍复用 `OpenFileOrActivateTab`，不会产生重复标签页
  - [ ] `dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRuleSessionPersistenceTests"` 通过

  **QA Scenarios**:
  ```
  Scenario: 会话 round-trip 后规则完整保留
    Tool: Bash
    Steps: 运行 dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRuleSessionPersistenceTests"
    Expected: XML/default-session 持久化回归测试 PASS
    Evidence: .sisyphus/evidence/task-5-monitor-session.txt

  Scenario: 命中已打开文件时激活旧标签页而不是再开新标签
    Tool: Bash
    Steps: 运行包含 open-dispatch helper 的回归测试；若 helper 被抽出则按 helper 测试过滤执行
    Expected: 断言只发生一次打开请求或只返回 activate 行为
    Evidence: .sisyphus/evidence/task-5-monitor-session-error.txt
  ```

  **Commit**: YES | Message: `feat(main): restore monitor rules with sessions` | Files: [`src/MainForm.cs`, `tests/SnakeTail.Tests/MonitorRuleSessionPersistenceTests.cs`]

- [ ] 6. 实现单条规则编辑对话框（MonitorRuleEditForm）

  **What to do**:
  - 新建 `src/MonitorRuleEditForm.cs`、`src/MonitorRuleEditForm.Designer.cs`、`src/MonitorRuleEditForm.resx`
  - 对话框布局固定为：
    - `Name` 文本框（可空）
    - `Directory` 文本框 + `Browse...` 按钮
    - `Pattern` 文本框（示例占位：`*err*.log`）
    - `Preview` 只读文本框，实时显示 `<Directory>\<Pattern>`
    - `Enabled` 复选框
    - `Preview Matches` 按钮 + 只读列表（最多展示前 20 条当前匹配文件，超出时显示“+N more”）
    - `OK` / `Cancel`
  - `Preview Matches` 只调用 `Directory.GetFiles(directory, pattern, TopDirectoryOnly)` 做同步预览；不创建 watcher
  - `OK` 前必须走 `MonitorRuleValidation`
  - `ShowDialog(owner)` 必须传 `MainForm` 或上级对话框 owner
  - 沿用当前项目的 WinForms 手写/Designer 风格，不引入 MVVM

  **Must NOT do**:
  - 不把规则编辑直接塞进 `TailConfigForm`
  - 不支持在该对话框里配置 per-rule 编码/颜色/窗口模式

  **Recommended Agent Profile**:
  - Category: `visual-engineering` - Reason: WinForms 交互与布局细节较多
  - Skills: [`my-ui`] - 需要遵守对话框、按钮、输入框与反馈规则
  - Omitted: [`superpowers:test-driven-development`] - 可测试逻辑已在 T2，UI 壳层以编译与最终 QA 为主

  **Parallelization**: Can Parallel: YES | Wave 2 | Blocks: [7] | Blocked By: [1, 2]

  **References**:
  - Pattern: `src/TailConfigForm.cs:25-35` - 现有配置对话框构造器模式
  - Pattern: `src/TailConfigForm.cs:61-188` - 载入/保存控件值的 WinForms 写法
  - Pattern: `src/MainForm.Designer.cs:128-166` - 菜单/对话框命名风格
  - API/Type: `src/MainForm.cs:1233-1237` - 对话框 owner 绑定示例
  - External: `README.md` - 当前产品已有“跟踪日志目录”能力，新增 UI 必须明确是“多文件规则管理”，不能与旧入口混淆

  **Acceptance Criteria**:
  - [ ] 编辑对话框可以新增/编辑一条 `MonitorRuleConfig`
  - [ ] Preview 文本随目录或 pattern 输入即时更新
  - [ ] `Preview Matches` 能显示当前目录前 20 个匹配文件或精确错误原因
  - [ ] 非法目录或 pattern 时，`OK` 不允许关闭且错误文案精确
  - [ ] `pwsh -File ".\script\build.ps1" --release` 编译通过

  **QA Scenarios**:
  ```
  Scenario: 编辑对话框对合法输入显示正确预览
    Tool: Bash
    Steps: 先运行 dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRuleValidationTests"，再运行 pwsh -File ".\script\build.ps1" --release
    Expected: 校验测试 PASS，发布构建成功，说明 UI 接线未破坏编译
    Evidence: .sisyphus/evidence/task-6-monitor-rule-edit.txt

  Scenario: 非法目录/模式不会被保存
    Tool: Bash
    Steps: 重跑校验测试中的失败用例，并执行发布构建
    Expected: 失败用例 PASS（即正确拦截），构建成功
    Evidence: .sisyphus/evidence/task-6-monitor-rule-edit-error.txt
  ```

  **Commit**: YES | Message: `feat(ui): add monitor rule edit dialog` | Files: [`src/MonitorRuleEditForm.cs`, `src/MonitorRuleEditForm.Designer.cs`, `src/MonitorRuleEditForm.resx`]

- [ ] 7. 实现规则管理总对话框（MonitorRulesForm）

  **What to do**:
  - 新建 `src/MonitorRulesForm.cs`、`src/MonitorRulesForm.Designer.cs`、`src/MonitorRulesForm.resx`
  - 对话框布局固定为：
    - 上半区 `ListView`：列为 `Enabled`、`Name`、`Directory`、`Pattern`、`Status`、`Matched`
    - 下半区 `ListView`：显示当前选中规则的已知匹配文件，列为 `FileName`、`FullPath`、`OpenState`、`LastSeen`
    - 右侧按钮：`Add...`、`Edit...`、`Delete`、`Enable/Disable`、`Refresh Now`、`Close`
  - 行为固定：
    - 选中规则变化时，下半区刷新该规则的匹配文件状态
    - 双击下半区文件时，调用 `MainForm.BringToFrontAndOpenOrActivateTab(new[]{path})`
    - `Refresh Now` 仅重跑当前规则的 baseline discover，不重建整个会话
    - 删除规则只删除规则，不关闭已打开标签页
  - 状态文案固定为：`Running` / `Disabled` / `Error: <reason>`
  - 若规则命中了已打开文件，下半区 `OpenState` 显示 `Open`；否则 `Pending/Open on match` 不需要，直接显示 `Managed`

  **Must NOT do**:
  - 不把 matched-files 列表做成树形结构
  - 不在删除规则时顺带关闭文件标签页

  **Recommended Agent Profile**:
  - Category: `visual-engineering` - Reason: 需要复杂一点的 WinForms 表格与交互反馈
  - Skills: [`my-ui`] - 列表、按钮、对话框 owner/ESC 行为需一致
  - Omitted: [`superpowers:test-driven-development`] - 主要依赖 T2/T4/T5 的可测逻辑与最终编译/QA

  **Parallelization**: Can Parallel: YES | Wave 2 | Blocks: [8] | Blocked By: [4, 6]

  **References**:
  - Pattern: `src/MainForm.cs:138-156` - 状态栏更新语义可借鉴到对话框状态文案刷新
  - Pattern: `src/MainForm.cs:546-617` - 双击匹配文件时必须复用既有打开/激活逻辑
  - Pattern: `src/MainForm.cs:655-673` - 右键/当前选择驱动菜单显示的现有风格
  - Pattern: `src/TailConfigForm.cs:250-257` - 列表项由配置对象 Tag 驱动的写法
  - API/Type: `src/Storage/SnakeTailStorage.cs:239-254` - 如需持久化对话框位置，可扩展 AppSettings 读写能力

  **Acceptance Criteria**:
  - [ ] 管理对话框能完整展示规则列表与当前匹配文件列表
  - [ ] 规则启用/停用、删除、Refresh Now 都能即时反映到状态列
  - [ ] 双击匹配文件时只会打开或激活一个标签页
  - [ ] `pwsh -File ".\script\build.ps1" --release` 编译通过

  **QA Scenarios**:
  ```
  Scenario: 规则列表与匹配文件列表状态能稳定刷新
    Tool: Bash
    Steps: 运行 pwsh -File ".\script\build.ps1" --release；再运行 dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRuleManagerTests"
    Expected: 构建成功，manager 状态测试 PASS，说明 UI 依赖的状态源稳定
    Evidence: .sisyphus/evidence/task-7-monitor-rules-form.txt

  Scenario: 删除/停用规则不会误关闭已打开标签页
    Tool: Bash
    Steps: 运行与 manager/会话相关的回归测试（MonitorRuleManagerTests + MonitorRuleSessionPersistenceTests）
    Expected: 测试 PASS，未出现“删除规则即关闭文件”的错误语义
    Evidence: .sisyphus/evidence/task-7-monitor-rules-form-error.txt
  ```

  **Commit**: YES | Message: `feat(ui): add monitor rules management dialog` | Files: [`src/MonitorRulesForm.cs`, `src/MonitorRulesForm.Designer.cs`, `src/MonitorRulesForm.resx`]

- [ ] 8. 接入 MainForm 菜单、对话框状态持久化与交互收尾

  **What to do**:
  - 修改 `src/MainForm.Designer.cs` 与 `src/MainForm.cs`
  - 在 `File` 菜单中新增项：`Monitor &Rules...`，位置固定在 `Open File...` 之后、`Open EventLog...` 之前
  - 点击后打开 `MonitorRulesForm`，owner 为 `this`
  - 对话框位置/大小持久化方案固定为：扩展 `SnakeTailStorage` 的 `AppSettings` 访问能力，新增以 key 形式保存 `MonitorRulesForm.Bounds` 与 `MonitorRuleEditForm.Bounds`
  - 若没有任何规则，管理对话框下半区显示空状态文案而不是空白异常
  - 当规则运行发生错误时，通过管理对话框状态列与必要日志反馈，不弹出阻塞式 message box 洪泛
  - 保持 tray 菜单与主菜单一致：由于现有托盘菜单会“借用” File 菜单项，新菜单无需额外复制实现，只要放入 File 菜单即可被复用

  **Must NOT do**:
  - 不新增第二套托盘专用实现
  - 不把运行时错误全部做成弹窗

  **Recommended Agent Profile**:
  - Category: `unspecified-high` - Reason: 涉及菜单结构、托盘复用、窗口状态持久化与错误反馈收口
  - Skills: [`my-ui`] - 需要遵守菜单、对话框 owner、一致性规则
  - Omitted: [`superpowers:test-driven-development`] - 重点是接线与状态持久化

  **Parallelization**: Can Parallel: YES | Wave 2 | Blocks: [9] | Blocked By: [5, 7]

  **References**:
  - Pattern: `src/MainForm.Designer.cs:128-166` - `File` 菜单定义位置
  - Pattern: `src/MainForm.cs:1233-1237` - `ShowDialog(this)` owner 口径
  - Pattern: `src/MainForm.cs:1258-1277` - 托盘菜单复用 `File` 菜单项的实现，新增菜单项必须兼容此机制
  - Pattern: `src/Storage/SnakeTailStorage.cs:239-254` - 可扩展为通用设置读写入口
  - Pattern: `src/MainForm.cs:1491-1558` - 窗口关闭时的资源释放与错误吞吐风格

  **Acceptance Criteria**:
  - [ ] `File` 菜单与 tray 菜单都能访问 `Monitor Rules...`
  - [ ] `MonitorRulesForm` 与 `MonitorRuleEditForm` 会记住上次位置和大小
  - [ ] 规则运行错误不会产生重复阻塞弹窗
  - [ ] `pwsh -File ".\script\build.ps1" --release` 编译通过

  **QA Scenarios**:
  ```
  Scenario: 主菜单与托盘菜单都含有同一入口
    Tool: Bash
    Steps: 运行 pwsh -File ".\script\build.ps1" --release
    Expected: 编译成功，MainForm 菜单与托盘复用代码无编译/资源错误
    Evidence: .sisyphus/evidence/task-8-mainform-menu.txt

  Scenario: 对话框状态持久化键缺失时也能正常打开
    Tool: Bash
    Steps: 运行与 storage 相关的新增回归测试（若 executor 添加），然后执行发布构建
    Expected: 回归测试 PASS；首次打开对话框使用默认位置且不抛异常
    Evidence: .sisyphus/evidence/task-8-mainform-menu-error.txt
  ```

  **Commit**: YES | Message: `feat(main): wire monitor rules dialogs` | Files: [`src/MainForm.cs`, `src/MainForm.Designer.cs`, `src/Storage/SnakeTailStorage.cs`]

- [ ] 9. 更新文档并补齐最终回归说明

  **What to do**:
  - 更新 `README.md`：
    - 新增“按目录 + 通配符监控多个文件”的功能说明
    - 明确首版限制：仅当前目录、仅 `*`/`?`、不递归、不支持排除规则
    - 写出界面入口：`File -> Monitor Rules...`
    - 写出与旧 `FileCheckPattern` 的区别：旧能力是“单个 Tail 按模式追踪最新文件”，新能力是“会话级规则管理多个匹配文件”
  - 更新 `CLAUDE.md`：补充该功能的限制与 README 同步要求（只改项目级 CLAUDE，不改用户级全局说明）
  - 在 README 中补充验证命令：`dotnet test ...` 与 `pwsh -File .\script\build.ps1 --release`

  **Must NOT do**:
  - 不省略旧能力与新能力的差异说明
  - 不把“未来可能支持递归/正则”写成已实现能力

  **Recommended Agent Profile**:
  - Category: `writing` - Reason: 纯文档整理与边界说明
  - Skills: [] - 项目文档即可
  - Omitted: [`my-ui`] - 无界面实现

  **Parallelization**: Can Parallel: YES | Wave 2 | Blocks: [F1, F2, F3, F4] | Blocked By: [5, 8]

  **References**:
  - Pattern: `README.md` - 现有功能列表与构建说明位置
  - Pattern: `script/build.ps1:76-102` - 发布构建命令写法
  - Pattern: `0run.ps1:90-110` - 项目内测试运行入口
  - API/Type: `src/LogFileStream.cs:220-237` - 旧通配模式“找最新文件”的实现证据

  **Acceptance Criteria**:
  - [ ] README 清楚区分旧通配 Tail 与新监控规则功能
  - [ ] README 写明限制、入口与验证命令
  - [ ] CLAUDE.md 同步新增功能约束说明

  **QA Scenarios**:
  ```
  Scenario: 文档与实际命令一致
    Tool: Bash
    Steps: 按 README 中新增命令执行 dotnet test "tests/SnakeTail.Tests/SnakeTail.Tests.csproj" --configuration Debug --filter "FullyQualifiedName~MonitorRule"，再执行 pwsh -File ".\script\build.ps1" --release
    Expected: README 中记录的命令可直接运行成功
    Evidence: .sisyphus/evidence/task-9-docs.txt

  Scenario: README 未错误宣称超出首版范围的能力
    Tool: Bash
    Steps: 审阅 README/CLAUDE 变更并执行构建命令
    Expected: 文档仅覆盖单层目录 + * / ? 语义；构建无异常
    Evidence: .sisyphus/evidence/task-9-docs-error.txt
  ```

  **Commit**: YES | Message: `docs: describe monitor rules feature` | Files: [`README.md`, `CLAUDE.md`]

## Final Verification Wave (MANDATORY — after ALL implementation tasks)
> 4 review agents run in PARALLEL. ALL must APPROVE. Present consolidated results to user and get explicit "okay" before completing.
> **Do NOT auto-proceed after verification. Wait for user's explicit approval before marking work complete.**
> **Never mark F1-F4 as checked before getting user's okay.** Rejection or user feedback -> fix -> re-run -> present again -> wait for okay.
- [ ] F1. Plan Compliance Audit — oracle
- [ ] F2. Code Quality Review — unspecified-high
- [ ] F3. Real Manual QA — unspecified-high (+ playwright if UI)
- [ ] F4. Scope Fidelity Check — deep

## Commit Strategy
- 按任务提交，避免把 DTO、运行时、UI、文档混成一个大提交
- 推荐提交顺序：
  1. `feat(config): add monitor rule session model`
  2. `feat(monitor): add rule validation helpers`
  3. `feat(monitor): add directory watcher`
  4. `feat(monitor): add monitor rule manager`
  5. `feat(main): restore monitor rules with sessions`
  6. `feat(ui): add monitor rule edit dialog`
  7. `feat(ui): add monitor rules management dialog`
  8. `feat(main): wire monitor rules dialogs`
  9. `docs: describe monitor rules feature`

## Success Criteria
- 用户可在 `MainForm` 通过 `File -> Monitor Rules...` 打开规则管理对话框
- 用户可新增一条规则：目录 + `*err*.log`，启用后立即接管当前目录内所有匹配文件
- 新文件落入目录后自动打开或激活相应 Tail 标签页
- 同一文件不会因为重复事件或多条规则产生重复标签页
- 保存 XML 会话 / 默认会话后，规则在下次启动或重新加载会话时可恢复
- 现有单文件 Tail、旧 `FileCheckPattern` 行为、构建脚本、测试项目均不被破坏
