# Wildcard File Monitor MVP – WinForms design document

Date: 2026-04-07
Scope: Single-directory, non-recursive wildcard monitoring with dynamic discovery and tailing

- Objective
  - Provide a minimal, reliable UI-driven monitor that watches a user-specified directory for files matching a wildcard pattern (e.g. d:\\xxx\\*err*.log)
  - Auto-add newly matched files to the tailing workflow, each in its own TailForm tab

- Key components
  - MonitorRuleConfig: Name, DirectoryPath, FilePattern, Enabled
  - MonitorDirectoryWatcher: baseline discovery + FileSystemWatcher for new files
  - MonitorRuleManager: coordinates multiple watchers, deduplicates tailed files, emits events to UI
  - TailConfig: extend (add MonitorRules, additive, backward-compatible)
  - MonitorRuleEditForm, MonitorRulesForm: UI scaffolding to manage rules
  - MainForm: hook to launch MonitorRulesForm via Menu

- Data flow
  - User defines rules via MonitorRulesForm -> MonitorRuleManager keeps state
  - When a rule is enabled, a MonitorDirectoryWatcher performs a baseline scan and subscribes to FileMatched
  - Each matched path is tailed; duplicates are prevented via a per-run dedupe set
  - TailForm tabs are opened for each tailed file (one file per tab, deduped)

- Verification plan
  - Unit tests: MonitorRuleConfigTests, MonitorDirectoryWatcherTests, MonitorRuleManagerTests
  - Integration: ensure MonitorRulesForm can instantiate and pass rules to MonitorRuleManager
  - UI: verify opening MonitorRulesForm from MainForm and adding/editing rules

- Risks & mitigation
  - Potential race between baseline and Created event. Mitigate by deduplicating with a global tailedFiles set.
  - Non-recursive only; no recursive watching for initial MVP – documented as MVP constraint.
