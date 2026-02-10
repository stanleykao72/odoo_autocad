# AutoCAD-Odoo Claude Code Framework

> 專為 **AutoCAD-Odoo C# WPF 整合桌面應用**設計的 Claude Code 開發框架

---

## 概述

本框架提供完整的 Claude Code 配置，用於規範和加速 AutoCAD-Odoo 整合桌面應用的開發。包含多代理協作系統、自動觸發技能、程式碼品質規則、安全防護 Hooks 與標準文件模板。

## 框架架構

```
autocad-odoo-claude-code/
├── README.md                       # 本文件
├── deploy.ps1                      # PowerShell 部署腳本
├── deploy.sh                       # Bash 部署腳本
│
├── claude-md/                      # 階層式 CLAUDE.md
│   ├── CLAUDE.md                   # 根層級：全域規則與架構
│   └── csharp/
│       └── CLAUDE.md               # C# 層級：WPF/MVVM 開發規範
│
├── agents/                         # 6 個專業化代理
│   ├── wpf_developer_agent.md      # WPF/MVVM 開發者
│   ├── autocad_integration_agent.md # AutoCAD COM 整合專家
│   ├── odoo_api_developer_agent.md  # Odoo REST API 開發者
│   ├── qa_tester_agent.md          # QA 測試工程師
│   ├── project_coordinator_agent.md # 專案協調者
│   └── security_reviewer_agent.md   # 安全審查員
│
├── skills/                         # 11 個自動觸發技能
│   ├── README.md
│   ├── csharp-mvvm/                # CommunityToolkit.Mvvm 模式
│   ├── xaml-patterns/              # WPF XAML 開發模式
│   ├── dotnet-build/               # .NET 建置與套件管理
│   ├── autocad-com/                # AutoCAD COM Interop
│   ├── odoo-rest-api/              # Odoo REST API 整合
│   ├── git-workflow/               # Git 操作規範
│   ├── wpf-navigation/             # 導航系統機制
│   ├── di-patterns/                # 依賴注入配置
│   ├── i18n-cjk/                   # CJK 國際化支援
│   ├── debug-patterns/             # 除錯診斷策略
│   └── python-legacy/              # Python 舊系統分析 ★
│
├── rules/                          # 7 條永遠生效的規則
│   ├── README.md
│   ├── security.md                 # 安全規範 (CRITICAL)
│   ├── coding-style.md             # 編碼風格 (HIGH)
│   ├── testing.md                  # 測試規範 (HIGH)
│   ├── git-workflow.md             # Git 規範 (HIGH)
│   ├── wpf-patterns.md             # WPF 模式 (MEDIUM)
│   ├── odoo-patterns.md            # Odoo 整合模式 (MEDIUM)
│   └── agents.md                   # 代理調度規則 (MEDIUM)
│
├── hooks/                          # Hook 自動化系統
│   ├── README.md
│   └── hooks.json                  # 8 個 Hooks (4 Pre + 4 Post)
│
├── commands/                       # 4 個工作流命令
│   ├── tdd.md                      # TDD 開發流程
│   ├── code-review.md              # 程式碼審查
│   ├── git-push.md                 # 安全推送
│   └── plan.md                     # 實作規劃
│
├── templates/                      # 文件模板系統
│   ├── documents/
│   │   ├── README.md               # 模板系統說明
│   │   ├── US_template.md          # 用戶故事
│   │   ├── TASK_BE_template.md     # 後端任務
│   │   ├── TASK_FE_template.md     # 前端任務
│   │   ├── TS_template.md          # 測試規格
│   │   ├── BUG_template.md         # 缺陷報告
│   │   └── FR_template.md          # 功能需求
│   └── shared_handoffs/
│       ├── active_handoffs/
│       │   └── handoff_queue.md    # 交接佇列
│       └── agent_status/
│           └── agent_availability.md # 代理狀態
│
└── scripts/                        # 輔助腳本
    └── hooks/                      # Hook 實作腳本
```

## 核心功能

### 1. 多代理協作系統

6 個專業化代理各司其職：

| 代理 | 職責 | Model |
|------|------|-------|
| WPF 開發者 | UI/XAML/ViewModel 開發 | sonnet |
| AutoCAD 整合專家 | COM Interop、繪圖操作 | sonnet |
| Odoo API 開發者 | REST API、資料同步 | sonnet |
| QA 測試工程師 | 測試設計與執行 | sonnet |
| 專案協調者 | 任務分派與管理 | opus |
| 安全審查員 | 安全審計 | opus |

### 2. 自動觸發技能

當對話中出現關鍵字時自動載入對應技能，例如：
- 提到 `ViewModel` → 載入 `csharp-mvvm` skill
- 提到 `XAML` → 載入 `xaml-patterns` skill
- 提到 `Python` 或 `舊程式` → 載入 `python-legacy` skill

### 3. 安全防護 Hooks

**PreToolUse（阻擋危險操作）**:
- 阻擋 `git push --force`、`git reset --hard`
- 警告 `git add .`
- 修改 `.csproj` 時提醒

**PostToolUse（品質檢查）**:
- 編輯 `.cs` 後檢查 async void、new HttpClient、硬編碼密碼
- 編輯 `.xaml` 後檢查事件處理和 CJK 字型

### 4. 文件模板

標準化的開發文件格式：US（用戶故事）、TASK-BE/FE（任務）、TS（測試規格）、BUG（缺陷報告）、FR（功能需求）

## 部署

### PowerShell（推薦）

```powershell
# 預覽部署（不實際執行）
.\deploy.ps1 -DryRun

# 備份現有設定並部署
.\deploy.ps1 -Backup

# 部署到指定目錄
.\deploy.ps1 -Target "C:\Users\me\.claude"
```

### Bash

```bash
# 預覽部署
./deploy.sh --dry-run

# 備份並部署
./deploy.sh --backup

# 部署到指定目錄
./deploy.sh --target /home/user/.claude
```

### 部署內容

| 元件 | 數量 | 目標目錄 |
|------|------|---------|
| Skills | 11 | `~/.claude/skills/` |
| Agents | 6 | `~/.claude/agents/` |
| Commands | 4 | `~/.claude/commands/` |
| Rules | 7 | `~/.claude/rules/` |
| Templates | 8 | `~/.claude/templates/` |
| Hooks | 8 | `~/.claude/settings.json` (合併) |

### CLAUDE.md 處理

Deploy 腳本**不會**自動覆蓋專案的 `CLAUDE.md`。請手動決定：
- `claude-md/CLAUDE.md` → 專案根目錄
- `claude-md/csharp/CLAUDE.md` → `csharp/` 目錄

## 設計原則

本框架遵循以下原則：

1. **DO NOT OVERDESIGN** — 只做被要求的事
2. **SDD (Specification-Driven Development)** — 規格驅動開發
3. **TDD (Test-Driven Development)** — 測試驅動開發
4. **安全優先** — CRITICAL 安全問題立即修復
5. **正體中文** — 文件使用正體中文，技術術語保留英文

## 參考

本框架參考 [odoo-claude-code](https://github.com/stanleykao72/odoo-claude-code) 的架構設計，針對 C# WPF AutoCAD-Odoo 整合專案進行了完整的客製化。
