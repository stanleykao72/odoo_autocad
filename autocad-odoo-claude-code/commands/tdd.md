# TDD 開發命令

## 核心原則

- **DO NOT OVERDESIGN** — 只做被要求的事
- **Red → Green → Refactor** — 嚴格遵循 TDD 循環
- **最小實作** — 用最少的程式碼使測試通過

---

## 執行流程

### Phase 1: 任務分析

1. 確認待開發功能或待修復缺陷
2. 識別涉及的檔案和模組
3. 確定測試範圍（單元/整合）

### Phase 2: 撰寫測試（Red）

1. 在 `tests/OdooAutoCAD.Integration.Tests/` 中建立或找到對應的測試類別
2. 撰寫描述期望行為的測試方法
3. 測試命名：`{Method}_{Scenario}_{ExpectedResult}`
4. 執行測試確認**失敗**：

```bash
"C:\Program Files\dotnet\dotnet.exe" test csharp/OdooAutoCADIntegration/tests/OdooAutoCAD.Integration.Tests/ --filter "FullyQualifiedName~{TestName}" --verbosity normal
```

**Checkpoint**: 測試必須失敗且原因正確

### Phase 3: 實作（Green）

1. 實作**最少的程式碼**使測試通過
2. 不要新增測試未要求的功能
3. 執行測試確認**通過**：

```bash
"C:\Program Files\dotnet\dotnet.exe" test csharp/OdooAutoCADIntegration/tests/OdooAutoCAD.Integration.Tests/ --verbosity normal
```

**Checkpoint**: 所有測試必須通過

### Phase 4: 重構（Refactor）

1. 改善程式碼結構（命名、提取方法等）
2. 確保測試仍然通過
3. 消除重複程式碼

**Checkpoint**: 所有測試仍然通過

### Phase 5: 驗證

1. 完整建置：
```bash
"C:\Program Files\dotnet\dotnet.exe" build csharp/OdooAutoCADIntegration/OdooAutoCADIntegration.sln
```

2. 全部測試通過
3. 無新增建置警告

### Phase 6: 提交

- 行為變更和結構變更**分開提交**
- Commit message 遵循格式規範

---

## 代理指派

| 任務類型 | 代理 |
|---------|------|
| ViewModel 測試 | QA 測試工程師 → WPF 開發者 |
| 服務測試 | QA 測試工程師 → 對應開發者 |
| COM 整合測試 | QA 測試工程師 → AutoCAD 整合專家 |
| API 測試 | QA 測試工程師 → Odoo API 開發者 |
