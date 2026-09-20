# 網路客戶詢問統計 App

手機版 PWA，專門整理承邦有限公司每月網路客戶詢問。

## 使用流程
1. 按「執行本月統計」
2. 取得 Outlook Connector 回傳的本月郵件
3. 篩選：寄件人或收件人為 sales@cbtrade.com.tw，或主旨包含「聯絡我們」
4. 解析 FW/RE 轉寄內容中的原始客戶資料
5. 排除垃圾、促銷、廣告與明顯內部郵件
6. 去除重複客戶
7. 顯示統計並匯出 Excel

## 安全
前端不保存 Microsoft access token。Outlook 取信應透過安全 Connector/OAuth 後端完成。