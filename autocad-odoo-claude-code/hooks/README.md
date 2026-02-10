# Hooks 系統

## 概述

Hooks 是自動化的工具攔截器，在 Claude Code 執行工具操作前後觸發，提供安全防護和品質檢查。

## Hook 類型

| 類型 | 說明 | 時機 |
|------|------|------|
| `PreToolUse` | 工具執行前攔截 | 可阻擋危險操作 |
| `PostToolUse` | 工具執行後檢查 | 品質驗證和提醒 |

## 目前的 Hooks

### PreToolUse（防護）

| Hook | 觸發條件 | 動作 |
|------|---------|------|
| 破壞性 git 命令阻擋 | `git push --force`, `git reset --hard`, `git checkout .` | 阻擋執行 |
| git add 警告 | `git add .`, `git add -A` | 顯示警告 |
| 非必要文件阻擋 | 建立 `.md`/`.txt`（非 README/CLAUDE/CHANGELOG） | 阻擋執行 |
| .csproj 修改提醒 | 編輯 `.csproj` 檔案 | 顯示注意事項 |

### PostToolUse（品質檢查）

| Hook | 觸發條件 | 動作 |
|------|---------|------|
| C# 品質檢查 | 編輯 `.cs` 檔案 | 檢查 async void, new HttpClient |
| XAML 品質檢查 | 編輯 `.xaml` 檔案 | 檢查 Click 事件、CJK 字型 |
| PR URL 記錄 | `gh pr create` | 記錄 PR 連結 |
| 安全掃描 | 編輯 `.cs` 檔案 | 偵測硬編碼秘密 |

## 部署

Hooks 配置需合併到 `~/.claude/settings.json` 中。使用 deploy 腳本自動處理。

## 新增 Hook

1. 在 `hooks.json` 的對應類型中加入新項目
2. 設定 `matcher`（匹配條件）
3. 定義 `hooks`（執行的命令）
4. 加入 `description`（說明）
5. 更新本 README
