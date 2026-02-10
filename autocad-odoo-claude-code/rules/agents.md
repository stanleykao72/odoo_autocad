# 代理調度規範

> **優先級**: MEDIUM
> **適用**: 專案協調者代理

---

## 1. 可用代理

| 代理 | 檔案 | 使用場景 |
|------|------|---------|
| WPF 開發者 | `wpf_developer_agent.md` | UI、XAML、ViewModel、樣式 |
| AutoCAD 整合專家 | `autocad_integration_agent.md` | COM Interop、繪圖操作 |
| Odoo API 開發者 | `odoo_api_developer_agent.md` | REST API、資料同步 |
| QA 測試工程師 | `qa_tester_agent.md` | 測試設計與執行 |
| 專案協調者 | `project_coordinator_agent.md` | 任務分派與管理 |
| 安全審查員 | `security_reviewer_agent.md` | 安全審計 |

## 2. 自動調度條件

| 條件 | 調度代理 |
|------|---------|
| 修改 `.xaml` 或 `ViewModels/` | WPF 開發者 |
| 涉及 COM、AutoCAD、Marshal | AutoCAD 整合專家 |
| 涉及 HttpClient、Odoo API | Odoo API 開發者 |
| 任務含「測試」或 TDD | QA 測試工程師 |
| 跨代理任務、Sprint 規劃 | 專案協調者 |
| PR 前審查、安全問題 | 安全審查員 |
| 涉及 Python 舊碼分析 | 對應代理 + python-legacy skill |

## 3. 平行執行規則

### 可平行
- WPF 開發者 + Odoo API 開發者（UI 和 API 無直接依賴）
- 多個測試任務

### 不可平行
- 服務介面設計 → 服務實作（順序依賴）
- 開發 → 測試（需先完成開發）
- 功能開發 → 安全審查（需先完成功能）

## 4. 交接協議

### 交接模板
```markdown
## 代理交接紀錄

| 欄位 | 值 |
|------|-----|
| 來源代理 | {agent} |
| 目標代理 | {agent} |
| 任務 ID | {task_id} |
| 日期 | {date} |

### 已完成
- {item 1}
- {item 2}

### 待處理
- {item 1}

### 相關檔案
- `path/to/file1.cs`
- `path/to/file2.xaml`

### 注意事項
- {note}
```

## 5. 衝突解決

當多個代理對同一檔案有修改需求時：
1. 專案協調者決定執行順序
2. 先完成的代理提交變更
3. 後執行的代理在最新程式碼上工作
4. 禁止平行修改同一檔案
