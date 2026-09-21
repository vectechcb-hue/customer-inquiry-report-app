# VECTECH LINE 網路客戶統計後端

本目錄是 LINE Official Account -> Webhook -> Cloudflare Worker -> D1 的後端範本。

必要的 Cloudflare Worker Secrets：
LINE_CHANNEL_SECRET：LINE Developers Console 的 Channel Secret
LINE_CHANNEL_ACCESS_TOKEN：Messaging API Channel Access Token
LINE_READ_API_KEY：供 VECTECH App 讀取 LINE 統計資料的獨立 API Key

請不要把上述秘密寫入 GitHub 原始碼。Channel Secret 用來驗證 LINE Webhook 的 x-line-signature。

LINE Webhook URL：
https://你的-worker.workers.dev/webhook

資料查詢 API：
https://你的-worker.workers.dev/messages?from=2026-09-01T00:00:00.000Z&to=2026-10-01T00:00:00.000Z

讀取 API 時使用 HTTP Header：
X-API-Key: 你的 LINE_READ_API_KEY

回傳資料會包含 eventId、userId、timestamp、messageType、text、displayName、salesperson，VECTECH App 會再做網路客戶判斷與統計。

部署步驟：
1. Cloudflare 建立 D1，名稱可用 vectech-line-customer-db。
2. 執行 line-worker/schema.sql。
3. 將 line-worker/worker.js 部署成 Cloudflare Worker。
4. 設定上述三個 Worker Secrets。
5. LINE Developers Console 開啟 Use webhook，填入 /webhook。
6. 將 Worker 網址貼到 VECTECH App 的 LINE 官方帳號設定，按測試。

安全性：LINE 官方文件要求 Webhook 在伺服器端用 Channel Secret 驗證 x-line-signature；Worker 已先驗證原始 request body，再處理事件。