# 拾逅花事｜進銷存系統

自行開發的花藝品牌 POS + 進銷存系統，基於 Google Apps Script 後端 + HTML 前端。

## 系統架構

- **Code.gs** — Google Apps Script 後端主程式（試算表 API、交易、會員、報表、進貨、報廢）
- **index.html** — POS 主結帳畫面
- **dashboard.html** — 營運儀表板
- **inventory.html** — 庫存管理
- **procurement.html** — 進貨管理
- **products.html** — 商品管理

## 支援的點位

- 高雄 FOCUS 13
- 台南 FOCUS
- 誠品生活台南
- Studio（Studio 訂單、婚禮捧花）

## 版本

目前版本：v2.2（2026-04）

Changelog 詳見 `Code.gs` 檔案頂部。

## 部署

1. Google Sheets 建立試算表，記下 ID
2. 「延伸功能」→「Apps Script」貼入 `Code.gs`
3. 設定 Script Properties 的 `SPREADSHEET_ID`
4. 執行 `setupSheets()` 與 `setupProcurementSheets()` 初始化
5. 部署為網頁應用程式（我自己執行 / 所有人可存取）
6. 前端 HTML 填入部署 URL 即可使用
