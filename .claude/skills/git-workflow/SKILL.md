# Git 工作流 Skill

---
name: git-workflow
description: Git 操作規範、commit 格式、分支管理
trigger-keywords:
  - git commit
  - 分支
  - merge
  - PR
  - pull request
  - branch
  - push
  - rebase
allowed-tools:
  - Bash
  - Read
---

## 概述

本 Skill 規範 Git 操作流程，確保版本控制的一致性與安全性。

## 絕對禁止

- ❌ `git add . && git commit && git push`（合併命令）
- ❌ `git add .` 或 `git add -A`（必須指定檔案）
- ❌ `git push --force`（除非明確要求）
- ❌ `git reset --hard`（除非明確要求）
- ❌ `git checkout .`（除非明確要求）
- ❌ 直接修改 commit message（不使用 `--amend`，除非要求）

## 標準工作流

### 1. 查看狀態
```bash
git status
git diff
```

### 2. 暫存指定檔案
```bash
git add csharp/OdooAutoCADIntegration/src/OdooAutoCAD.App/ViewModels/MainViewModel.cs
git add csharp/OdooAutoCADIntegration/src/OdooAutoCAD.App/Views/Pages/DashboardPage.xaml
```

### 3. 提交（使用 HEREDOC）
```bash
git commit -m "$(cat <<'EOF'
feat: Add dashboard page with navigation support

- Implement DashboardViewModel with data loading
- Create DashboardPage.xaml with CJK font support
- Wire up navigation from MainViewModel

Co-Authored-By: Claude Opus 4.6 <noreply@anthropic.com>
EOF
)"
```

### 4. 推送（需確認）
```bash
git push origin csharp.0.0.1
```

## Commit Message 格式

```
<type>: <description>

<body>

Co-Authored-By: Claude Opus 4.6 <noreply@anthropic.com>
```

### Type 分類

| Type | 說明 | 範例 |
|------|------|------|
| `feat` | 新功能 | `feat: Add AutoCAD connection service` |
| `fix` | 修復缺陷 | `fix: Resolve navigation crash on empty page` |
| `refactor` | 重構 | `refactor: Extract navigation logic to service` |
| `test` | 測試 | `test: Add shutdown cleanup tests` |
| `docs` | 文件 | `docs: Update CLAUDE.md build instructions` |
| `style` | 格式 | `style: Fix XAML indentation` |
| `chore` | 雜務 | `chore: Update NuGet packages` |

## 分支策略

| 分支 | 用途 |
|------|------|
| `master` | 主分支（PR 目標） |
| `csharp.x.y.z` | C# 開發分支 |
| `feature/*` | 功能分支 |
| `fix/*` | 修復分支 |

## 檢查清單

- [ ] `git status` 確認變更檔案
- [ ] 使用具體檔案名稱 `git add`
- [ ] commit message 遵循格式
- [ ] 包含 `Co-Authored-By` 行
- [ ] 推送前確認目標分支正確
