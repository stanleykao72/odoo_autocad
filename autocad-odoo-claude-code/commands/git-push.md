# Git Push 命令

## 核心原則

- **DO NOT OVERDESIGN** — 只做必要的 git 操作
- 安全第一 — 推送前必須確認

---

## 執行流程

### Step 1: 環境檢查

```bash
# 確認當前分支
git branch --show-current

# 確認遠端追蹤
git remote -v

# 確認狀態（無未提交變更）
git status
```

**Checkpoint**: 無未暫存或未提交的變更

### Step 2: 變更偵測

```bash
# 查看待推送的 commits
git log origin/{branch}..HEAD --oneline

# 查看完整差異
git diff origin/{branch}..HEAD --stat
```

**Checkpoint**: 確認待推送的 commits 正確

### Step 3: 建置驗證

```bash
# 確認建置通過
"C:\Program Files\dotnet\dotnet.exe" build csharp/OdooAutoCADIntegration/OdooAutoCADIntegration.sln

# 確認測試通過
"C:\Program Files\dotnet\dotnet.exe" test csharp/OdooAutoCADIntegration/tests/OdooAutoCAD.Integration.Tests/ --verbosity normal
```

**Checkpoint**: 建置和測試都通過

### Step 4: 推送

```bash
git push origin {branch}
```

### 安全檢查

推送前自動驗證：
- [ ] 不是 `--force` 推送
- [ ] 目標不是 `master`/`main`（除非是 PR merge）
- [ ] 無敏感檔案在提交中
- [ ] 建置通過
- [ ] 測試通過
