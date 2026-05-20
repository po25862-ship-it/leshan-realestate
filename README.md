# 捷運樂善店 業務助手

## 部署步驟

### 1. 上傳到 GitHub
1. 去 github.com 建立新 repository（名稱：leshan-realestate）
2. 把這個資料夾所有檔案上傳

### 2. 部署到 Vercel
1. 去 vercel.com，用 GitHub 登入
2. 點「Add New Project」→ 選你的 repo
3. Framework 選「Create React App」
4. 點「Deploy」
5. 完成後會得到網址，例如：https://leshan-realestate.vercel.app

### 3. 加到手機主畫面
- iPhone：Safari 打開網址 → 分享 → 加入主畫面
- Android：Chrome 打開網址 → 選單 → 加入主畫面

## 資料架構
| 資料 | 儲存 |
|------|------|
| 物件庫 | Firebase（共用）|
| 買方資料 | Firebase（共用）|
| 帶看記錄 | Firebase（共用）|
| 行事曆 | Firebase（共用）|
| 電話/Line | 各自手機（私人）|
