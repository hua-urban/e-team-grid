# E 組九宮格展示網站

BNI 華都更分會 E 組的九宮格網站，網址：https://hua-urban.com/e-team-grid/

## 檔案分工

- `index.html`、`data.js`、`images/`：網站本體，由 GitHub Pages 從 `main` 分支根目錄發布。
- `backend/`：Apps Script 後端原始碼（`Code.gs`）與部署筆記。**不會被發布到網站**——`_config.yml` 的 `exclude` 把它排除在 Pages 建置之外（2026-10-21 起）。
  後端實際跑的是綁在「E組九宮格_後台資料」試算表上的 Apps Script，改版步驟見 `backend/部署SOP.md`。

## 注意

- 這個 repo 是公開的。信箱、Telegram token 等設定值一律放 Apps Script 的「指令碼屬性」，不寫進任何檔案。
- 新增不該公開的資料夾或檔案時，記得一併加進 `_config.yml` 的 `exclude`。
