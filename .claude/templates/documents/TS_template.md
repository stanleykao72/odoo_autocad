# TS-{xxx}-{yy}: {測試規格標題}

> **狀態**: 草稿 | 審查中 | 已批准 | 執行中 | 已完成
> **指派代理**: QA 測試工程師
> **來源用戶故事**: US-{xxx}-{yy}
> **建立日期**: {YYYY-MM-DD}

---

## 測試範圍

{描述此測試規格涵蓋的功能範圍}

## 測試環境

| 項目 | 規格 |
|------|------|
| 框架 | xUnit |
| Mock | Moq |
| 建置工具 | dotnet test |
| 測試專案 | `OdooAutoCAD.Integration.Tests` |

## 測試案例

### 正向測試 (Happy Path)

| 編號 | 測試方法名稱 | 前置條件 | 操作 | 期望結果 |
|------|-------------|---------|------|---------|
| TC-01 | `{Method}_{Scenario}_{Expected}` | {前置條件} | {操作} | {期望結果} |
| TC-02 | `{Method}_{Scenario}_{Expected}` | {前置條件} | {操作} | {期望結果} |

### 反向測試 (Error Path)

| 編號 | 測試方法名稱 | 前置條件 | 操作 | 期望結果 |
|------|-------------|---------|------|---------|
| TC-E01 | `{Method}_{ErrorScenario}_{Expected}` | {前置條件} | {錯誤操作} | {錯誤處理} |

### 邊界測試 (Boundary)

| 編號 | 測試方法名稱 | 邊界條件 | 期望結果 |
|------|-------------|---------|---------|
| TC-B01 | `{Method}_{BoundaryCase}_{Expected}` | {邊界條件} | {期望結果} |

### 整合測試

| 編號 | 測試方法名稱 | 涉及元件 | 期望結果 |
|------|-------------|---------|---------|
| TC-I01 | `{Flow}_{Scenario}_{Expected}` | {元件列表} | {期望結果} |

## Mock 設定

```csharp
// 需要 Mock 的外部依賴
var mockService = new Mock<I{Name}Service>();
mockService.Setup(s => s.{Method}({params}))
           .ReturnsAsync({expected_value});
```

## 測試資料

| 資料集 | 用途 | 值 |
|--------|------|-----|
| {data_1} | {用途} | {值} |

## 執行指令

```bash
# 執行此測試規格的所有測試
"C:\Program Files\dotnet\dotnet.exe" test --filter "FullyQualifiedName~{TestClassName}" --verbosity normal
```

## 驗收標準

- [ ] 所有測試案例已實作
- [ ] 所有測試通過
- [ ] 覆蓋 happy path 和 error path
- [ ] 外部依賴已 Mock
- [ ] 無相互依賴的測試
