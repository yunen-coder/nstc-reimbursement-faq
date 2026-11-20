國科會計畫經費報銷查詢網站 - Cloudflare Pages 部署包

資料夾內容
- index.html：原本網站畫面，已將 google.script.run 改為 /api/* 呼叫
- functions/api/[[path]].js：Cloudflare Pages API 代理
- apps-script/Code.gs：需貼到原本綁定 Google 試算表的 Apps Script 專案

一、設定 Google Apps Script
1. 開啟原本綁定資料試算表的 Apps Script 專案。
2. 用本資料夾 apps-script/Code.gs 全文取代原本程式碼並儲存。
3. 執行 setupSheets() 一次，依提示授權。
4. 部署 > 新增部署作業 > 類型選「網頁應用程式」。
5. 執行身分選「我」；誰可以存取選「任何人」。
6. 部署後複製 /exec 結尾的網頁應用程式網址。若日後修改程式，請管理部署並更新版本。

二、部署 Cloudflare Pages
1. 將 cloudflare_faq 資料夾內容上傳到 GitHub repository（根目錄需直接看見 index.html 和 functions 資料夾）。
2. Cloudflare 控制台 > Workers & Pages > Create application > Pages > Connect to Git。
3. 選擇該 GitHub repository。
4. Framework preset 選 None；Build command 留空；Build output directory 填 /（或依控制台要求填 .）。
5. 建立網站後，進入 Pages 專案 > Settings > Variables and Secrets。
6. 新增環境變數 GAS_URL，值填第一步複製的 Apps Script /exec 網址；Production 與 Preview 若都要測試，兩個環境都設定。
7. 重新部署 Pages。完成後可使用 Cloudflare 提供的 *.pages.dev 網址。

三、重要提醒
- 此版本由 Cloudflare Pages Function 代為呼叫 Apps Script，避免瀏覽器直接跨網域呼叫造成 CORS 問題。
- Apps Script「任何人」代表知道 API 網址的人可呼叫讀取資料及記錄點閱。請確認試算表內的承辦人姓名、分機、Email、流程與內部資料適合公開；若只限校內使用，需另做登入/存取控管，不能直接照此公開部署。
- 目前原始後端程式中的新增資料/管理者密碼功能只有註解，沒有實際實作；本部署包保留既有的查詢、流程、網站、承辦人及熱門點閱功能，不提供網頁端編輯資料。
- 流程圖若使用 Google Drive 圖片，圖片分享權限需允許網站訪客檢視。
- 請先以測試資料確認網站可讀取，再正式分享網址。
