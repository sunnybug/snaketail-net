# 极简目录通配符监控入口

## TL;DR
> **Summary**: 在 File 菜单新增一个极简的“监控目录（通配符）”入口，只选目录、填通配符，立即开始监控，目录内新文件/变更文件自动追加到界面显示。
> **Deliverables**: SimpleDirectoryMonitorDialog.cs, 菜单入口新增, 监控逻辑对接, README 补充说明
> **Effort**: Short
> **Parallel**: NO
> **Critical Path**: 创建对话框 → 新增菜单 → 对接监控并自动打开Tail → 更新文档 → 构建

## Context
### Original Request
“这个功能太复杂了，我只需要能监控指定目录下模糊匹配的文件，若有旧文件内容变动或新建文件，能即时监控到并输出到工具界面上”

### 现有情况
- 已有 MonitorDirectoryWatcher：单个目录非递归监控，支持通配符 *?，能发现新文件
- 已有 TailForm：可打开单个文件并实时尾追
- 已加的 Open Wildcard Monitor... 入口：会打开复杂配置对话框
- 期望：极简两步：选目录 + 填通配符 → 立即开始监控，新文件/变更文件自动显示

## Work Objectives
### Core Objective
新增菜单入口，直接让用户选目录、填通配符，确定后立即对该目录开始通配符监控，目录下匹配的新文件/变更文件自动追加到一个 MDI 子窗口中显示。

### Deliverables
1. 极简对话框 `SimpleDirectoryMonitorDialog.cs`：只选目录 + 填通配符（默认 *.log）
2. File 菜单新增第二个入口：“监控目录（通配符）”（或“Watch Directory (Wildcard)”）
3. 监控逻辑：选择后创建 MonitorDirectoryWatcher，匹配的文件自动用 TailForm 打开并显示到界面
4. README.md 更新：说明两个入口的区别
5. 构建验证通过

### Definition of Done
- Release 构建通过（./script/build.ps1 --release）
- 菜单可见：File → 监控目录（通配符）
- 点击后弹出对话框，可选目录、填通配符
- 确定后目录内匹配的已有/新文件自动在界面上用 TailForm 打开并尾追

### Must Have
- 极简 UI：没有复杂配置，只有目录 + 通配符
- 自动监控并打开新文件
- 复用 MonitorDirectoryWatcher 和 TailForm，不重复造轮子

### Must NOT Have
- 不要打开复杂的 TailConfigForm
- 不要要求用户配置编码/间隔/颜色等
- 不修改现有逻辑，只新增最小化代码

## Verification Strategy
- 构建通过 + 测试通过
- QA 手动走菜单：打开对话框 → 选本地日志目录 → 确定 → 看新文件/变更文件是否显示
- 证据：截图或录屏（可选）

## Execution Strategy
### Steps
1. 创建 SimpleDirectoryMonitorDialog.cs：只含目录选择 + 通配符文本框
2. 在 MainForm.Designer.cs 的 File 菜单下新增第二个入口（放在 Open File 下面）
3. 在 MainForm.cs 中实现菜单点击逻辑：打开对话框 → 创建 MonitorDirectoryWatcher → 对匹配文件调用 OpenFileSelection 或单独打开 TailForm
4. 更新 README.md 说明两个入口：“极简监控目录” vs “完整通配符配置”
5. 运行构建 ./script/build.ps1 --release，确保 0 错误

## TODOs
- [ ] 1. 创建 SimpleDirectoryMonitorDialog.cs（极简选目录+通配符）

  **What to do**: 在 src/ 下创建 SimpleDirectoryMonitorDialog.cs，只包含：
  - Label + TextBox + 按钮（浏览目录）
  - Label + TextBox（通配符，默认 *.log）
  - 确定/取消按钮
  - 验证：目录存在，通配符非空

  **Must NOT do**: 不要有任何其他配置选项（编码/颜色/间隔等）

  **References**: 可参考 OpenEventLogDialog.cs 的结构

  **Acceptance Criteria**: 可编译，对话框能弹出，能选目录，通配符可编辑

  **QA Scenarios**:
  ```
  Scenario: 打开对话框
    Tool: 手动（或后续通过 Playwright）
    Steps: 点击 File 菜单下的新入口
    Expected: 弹出 SimpleDirectoryMonitorDialog，默认通配符为 *.log
    Evidence: 截图或录屏片段
  ```

  **Commit**: YES | Message: "feat: add SimpleDirectoryMonitorDialog for minimal watch directory UI" | Files: [src/SimpleDirectoryMonitorDialog.cs]

- [ ] 2. 在 File 菜单新增“监控目录（通配符）”入口

  **What to do**: 在 MainForm.Designer.cs 的 File 菜单下新增菜单项，放在 openToolStripMenuItem 之后、openEventLogToolStripMenuItem 之前。文本：“Watch &Directory...” 或 “监控目录（通配符）”，快捷键可选 Ctrl+Shift+D 或不设快捷键。

  **Must NOT do**: 不要修改现有菜单项的顺序/大小/文本（除了必要的 Size 对齐）

  **References**: MainForm.Designer.cs 中 openToolStripMenuItem 的声明方式

  **Acceptance Criteria**: 菜单可见，位置在 Open File 下方

  **QA Scenarios**:
  ```
  Scenario: 菜单可见
    Tool: 手动
    Steps: 启动程序，打开 File 菜单
    Expected: 看到“监控目录（通配符）”（或“Watch Directory...”）
    Evidence: 截图
  ```

  **Commit**: YES | Message: "feat: add menu entry for simple directory monitor" | Files: [src/MainForm.Designer.cs]

- [ ] 3. 实现菜单点击逻辑：选目录后开始监控并自动打开匹配文件

  **What to do**: 在 MainForm.cs 中新增点击事件处理：
  - 打开 SimpleDirectoryMonitorDialog
  - 用户确定后创建 MonitorDirectoryWatcher
  - 对初始匹配的已有文件，调用 OpenFileSelection 或单独打开 TailForm
  - 对后续新增/变更的匹配文件，同样打开 TailForm 追加显示
  - 尽量复用 EnsureDefaultTailConfig() 和 CloneTailFileConfig() 避免重复代码

  **Must NOT do**: 不要打开 TailConfigForm，不要让用户再配置

  **References**: MonitorDirectoryWatcher.cs 的 FileMatched 事件、MainForm.cs 中 OpenFileSelection 的打开逻辑

  **Acceptance Criteria**: 选择目录后，目录下匹配通配符的已有/新文件能在界面上自动打开并尾追

  **QA Scenarios**:
  ```
  Scenario: 打开已有匹配文件
    Tool: 手动
    Steps: 在测试目录下放 a.log / b.log，选该目录，确定
    Expected: 两个文件都自动在界面上打开，开始尾追
    Evidence: 截图
  ```

  ```
  Scenario: 新增匹配文件自动打开
    Tool: 手动
    Steps: 打开监控后，在目录下新增 c.log
    Expected: c.log 自动在界面上打开并开始尾追
    Evidence: 截图
  ```

  **Commit**: YES | Message: "feat: implement simple directory monitor logic that auto-tails matched files" | Files: [src/MainForm.cs]

- [ ] 4. 更新 README.md，说明两个入口区别

  **What to do**: 在 README 的“跟踪日志目录...”条目下，补充说明：
  - File → 监控目录（通配符）：极简，只选目录+通配符，新文件自动打开
  - File → Open Wildcard Monitor...：完整配置，可配置编码/间隔/颜色等

  **References**: README.md 中现有的通配符说明

  **Acceptance Criteria**: README 中明确说明两个入口

  **Commit**: YES | Message: "docs: update README to explain both wildcard monitor entries" | Files: [README.md]

- [ ] 5. 最终验证：运行 Release 构建 + 测试

  **What to do**: 运行 ./script/build.ps1 --release，确保 0 错误/警告（忽略既有警告）。运行 dotnet test，确保 0 失败。

  **References**: CLAUDE.md

  **Acceptance Criteria**: 构建通过，测试通过

  **QA Scenarios**:
  ```
  Scenario: Release 构建通过
    Tool: Bash
    Steps: ./script/build.ps1 --release
    Expected: 0 错误，0 警告（忽略既有）
    Evidence: 构建输出
  ```

  **Commit**: NO | 不自动提交，由后续 /start-work 收尾时按约定提交

## Final Verification Wave
- [ ] F1. Plan Compliance Audit — 人工检查是否符合计划
- [ ] F2. 手动 QA 走通：菜单 → 对话框 → 选目录 → 自动打开新文件
- [ ] F3. Scope Fidelity Check — 确认没有加复杂配置，只有目录+通配符

## Commit Strategy
- 每个任务一个提交（除了 F 波）
- 按任务 1→2→3→4 顺序提交
- 最终由 /start-work 统一收尾提交（按 CLAUDE.md）

## Success Criteria
- 极简入口可见可用
- 选目录+通配符后，新文件/变更文件自动在界面显示
- Release 构建通过
