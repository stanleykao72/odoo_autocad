# Git 工作流規範

> **優先級**: HIGH
> **適用**: 所有代理

---

## 1. 絕對禁止

| 操作 | 原因 |
|------|------|
| `git add .` / `git add -A` | 可能包含敏感檔案或不相關變更 |
| `git push --force` | 可能覆蓋他人工作 |
| `git reset --hard` | 可能丟失未提交的工作 |
| `git checkout .` | 可能丟失未暫存的變更 |
| 合併命令 `git add . && git commit && git push` | 無法逐步驗證 |
| `--no-verify` | 跳過 pre-commit hooks |

## 2. Commit Message 格式

```
<type>: <description>

<optional body>

Co-Authored-By: Claude Opus 4.6 <noreply@anthropic.com>
```

### Type 分類
| Type | 說明 |
|------|------|
| `feat` | 新功能 |
| `fix` | 缺陷修復 |
| `refactor` | 重構（不改變行為） |
| `test` | 新增或修改測試 |
| `docs` | 文件更新 |
| `style` | 格式調整（不影響邏輯） |
| `chore` | 雜務（建置、套件更新） |

### Commit 指令格式
```bash
git commit -m "$(cat <<'EOF'
feat: Add dashboard page with data binding

- Create DashboardPage.xaml with CJK font support
- Implement DashboardViewModel with LoadDataCommand
- Register services in DI container

Co-Authored-By: Claude Opus 4.6 <noreply@anthropic.com>
EOF
)"
```

## 3. 標準操作流程

### 提交流程
```bash
# 1. 檢查狀態
git status

# 2. 檢查差異
git diff

# 3. 暫存特定檔案
git add path/to/file1.cs
git add path/to/file2.xaml

# 4. 提交
git commit -m "$(cat <<'EOF'
...message...
EOF
)"

# 5. 推送（需確認）
git push origin branch-name
```

### 分支策略
| 分支 | 用途 | 保護 |
|------|------|------|
| `master` | 主分支 | 受保護，需 PR |
| `csharp.x.y.z` | 開發分支 | 當前工作 |
| `feature/*` | 功能分支 | 短期 |
| `fix/*` | 修復分支 | 短期 |

## 4. 安全規則

- 提交前確認不包含敏感檔案（`.env`, `secrets.json`, `*.pfx`）
- 大型檔案（>1MB）需確認是否必要
- 二進位檔案（圖片、DLL）需確認是否應加入
- 合併衝突需解決後再提交（不刪除他人變更）
