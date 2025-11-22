```markdown
YTupload (exceljs) - 自動上傳並排程 YouTube 短片（Node.js 範例）

簡介
這個 CLI 程式會從本機 Excel (.xlsx) 的指定工作表讀取每支影片的 metadata（欄位：SEO_Title_ZH, SEO_Title_EN, Description_ZH, Description_EN, YT_Tags, Hashtags, 編號），依序上傳到 YouTube（可排程 publishAt），並把上傳結果（時間與 videoId）寫回 Excel 的「Uploaded」欄位。

重點變更：為避免 xlsx（SheetJS）已知漏洞，本版本改用 exceljs。

必要條件
- Node.js 16+
- 把 Google OAuth credentials（Desktop app）下載為 credentials.json 並放到專案根目錄（若你要自動上傳到 YouTube）
- 在 Google Cloud Console 啟用 YouTube Data API v3
- 將影片檔放在某資料夾，檔名以工作表「編號」為 base（例如 123.mp4 或 123_mobile.mp4）

安裝
1. 把檔案放入資料夾
2. npm install

執行
1. npm start 或 node index.js
2. 互動式輸入：
   - Excel 檔案路徑（.xlsx）
   - worksheet 名稱（預設 Output_100）
   - 編號欄位名稱（預設 編號）
   - 影片資料夾路徑
   - 是否啟用排程（若啟用請輸入起始時間與間隔分鐘）
3. 若是第一次執行會開啟瀏覽器讓你授權 Google 帳號（OAuth），完成授權後貼回授權碼或按流程完成（token.json 將被建立）。

欄位需求（工作表）
- 第一列為 header，下面每列為一支影片
- 需含（不區分大小寫）：
  - SEO_Title_ZH, SEO_Title_EN, Description_ZH, Description_EN, YT_Tags, Hashtags, 編號（或你指定的欄位名稱）
- 上傳成功後會在工作表新增/更新 Uploaded 欄位，格式為： 2025-11-22T03:00:00.000Z | videoId: abc123

安全說明
- 已從 xlsx 換成 exceljs，以避開 SheetJS 的已知高風險漏洞。
- credentials.json 與 token.json 請勿推到公開倉庫（請加入 .gitignore）。

若要變更
- 想要僅產生上傳清單（不上傳）或加入更複雜的檔名匹配，回覆我我會幫你調整。
```