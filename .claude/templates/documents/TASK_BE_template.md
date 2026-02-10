# TASK-BE-{xxx}-{yy}: {後端任務標題}

> **狀態**: PENDING | IN_PROGRESS | REVIEW | TESTING | DONE | BLOCKED
> **指派代理**: {agent_name}
> **來源用戶故事**: US-{xxx}-{yy}
> **建立日期**: {YYYY-MM-DD}

---

## 任務描述

{詳細描述此任務要實現的功能}

## 技術規格

### 涉及檔案

| 檔案 | 動作 | 說明 |
|------|------|------|
| `{path/to/file.cs}` | 新增/修改 | {說明} |

### 介面設計

```csharp
// 介面定義
public interface I{Name}Service
{
    {method signatures}
}
```

### 實作要點

1. {要點 1}
2. {要點 2}
3. {要點 3}

## DI 配置

```csharp
// 需要在 App.xaml.cs 中新增的註冊
services.AddSingleton<I{Name}Service, {Name}Service>();
```

## 資料模型

```csharp
public class {Name}Model
{
    {properties}
}
```

## 錯誤處理

| 錯誤情境 | 處理方式 |
|---------|---------|
| {情境 1} | {處理方式} |
| {情境 2} | {處理方式} |

## 安全考量

- {安全要點 1}
- {安全要點 2}

## 依賴

| 依賴項 | 狀態 |
|--------|------|
| {依賴 1} | 已完成/待處理 |

## 交付標準

- [ ] 程式碼實作完成
- [ ] 建置無新增錯誤/警告
- [ ] 單元測試通過
- [ ] 遵循 coding-style 規範
- [ ] 安全規範檢查通過
- [ ] 準備交接給 QA

## 交接紀錄

| 欄位 | 值 |
|------|-----|
| 完成日期 | {date} |
| 交接至 | {agent} |
| 備註 | {notes} |
