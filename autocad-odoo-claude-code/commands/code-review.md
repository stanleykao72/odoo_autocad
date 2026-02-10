# 程式碼審查命令

## 用途

對未提交的變更進行全面品質審查。

---

## 審查流程

### Step 1: 收集變更

```bash
git status
git diff
git diff --staged
```

### Step 2: 分類審查

#### 🔴 CRITICAL — 安全問題（立即修復）

| 項目 | 檢查方式 |
|------|---------|
| 硬編碼密碼/Key | `Grep: pattern="password\s*=\s*\"" type="cs"` |
| COM 物件未釋放 | 檢查 `Marshal.GetActiveObject` 對應的 `ReleaseComObject` |
| SQL 注入 | 檢查字串拼接 SQL |
| 路徑遍歷 | 檢查未驗證的檔案路徑 |

#### 🟠 HIGH — 功能影響（提交前修復）

| 項目 | 檢查方式 |
|------|---------|
| async void | `Grep: pattern="async\s+void" type="cs"` |
| 空 catch 區塊 | `Grep: pattern="catch\s*\([^)]*\)\s*\{\s*\}" type="cs"` |
| new HttpClient | `Grep: pattern="new\s+HttpClient" type="cs"` |
| 遺漏的 DI 註冊 | 比對新增的服務與 App.xaml.cs |
| 遺漏的測試 | 新增功能是否有對應測試 |

#### 🟡 MEDIUM — 品質問題（應處理）

| 項目 | 檢查方式 |
|------|---------|
| 命名不規範 | 檢查命名是否遵循 coding-style 規範 |
| CJK 字型遺漏 | XAML 中中文文字未使用 CJK 字型鏈 |
| 硬編碼字串 | XAML 中使用硬編碼值而非 StaticResource |
| Code-behind 邏輯 | `.xaml.cs` 中包含業務邏輯 |

#### 🟢 LOW — 建議改善

| 項目 | 說明 |
|------|------|
| using 排列 | 是否按規範分組排列 |
| 方法長度 | 超過 30 行的方法 |
| 註解品質 | 註解是否有價值 |

### Step 3: 生成報告

```markdown
## 程式碼審查報告

### 審查範圍
- 修改檔案: {count}
- 新增檔案: {count}
- 刪除檔案: {count}

### 發現問題

#### 🔴 CRITICAL ({count})
- [{file}:{line}] {description}

#### 🟠 HIGH ({count})
- [{file}:{line}] {description}

#### 🟡 MEDIUM ({count})
- [{file}:{line}] {description}

#### 🟢 LOW ({count})
- [{file}:{line}] {description}

### 建議
{recommendations}
```

---

## 代理指派

- **安全問題**: 安全審查員
- **程式碼品質**: 對應的開發代理
- **測試覆蓋**: QA 測試工程師
