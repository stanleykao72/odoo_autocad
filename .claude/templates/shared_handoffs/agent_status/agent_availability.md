# 代理狀態追蹤

## 可用性矩陣

| 代理 | 狀態 | 當前任務 | Context 使用量 |
|------|------|---------|---------------|
| WPF 開發者 | AVAILABLE / BUSY | {task_id} | FRESH / MODERATE / HIGH |
| AutoCAD 整合專家 | AVAILABLE / BUSY | {task_id} | FRESH / MODERATE / HIGH |
| Odoo API 開發者 | AVAILABLE / BUSY | {task_id} | FRESH / MODERATE / HIGH |
| QA 測試工程師 | AVAILABLE / BUSY | {task_id} | FRESH / MODERATE / HIGH |
| 專案協調者 | AVAILABLE / BUSY | {task_id} | FRESH / MODERATE / HIGH |
| 安全審查員 | AVAILABLE / BUSY | {task_id} | FRESH / MODERATE / HIGH |

## Context 使用量定義

| 等級 | 說明 | 建議動作 |
|------|------|---------|
| FRESH | < 30% context | 可接受新任務 |
| MODERATE | 30-60% context | 可繼續當前任務 |
| HIGH | 60-85% context | 建議儘快完成當前任務 |
| CRITICAL | > 85% context | 需要交接或壓縮 context |

## 狀態更新規則

1. 代理開始新任務時更新為 BUSY
2. 任務完成或交接後更新為 AVAILABLE
3. Context 使用量達到 HIGH 時通知專案協調者
4. Context 使用量達到 CRITICAL 時強制交接
