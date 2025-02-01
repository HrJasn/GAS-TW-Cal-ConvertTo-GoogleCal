# Google Apps Script 取得人事行政總處網頁辦公日曆表並更新到 Google 日曆
 ## 功能
	1. 自動下載中華民國人事行政總處公開網頁的日曆表xls檔。
	2. 將xls檔轉換為 Google 試算表然後將表格內容解析為條列式行程。
	3. 條列式行程排除週一到週五的上班日和週六、日的休假日，日期有節日名稱除外。
	4. 更新或新增排除後的條列式行程寫入到 Apps Script 使用的試算表。
	5. 更新或新增排除後的條列式行程寫入到 Apps Script 使用試算表名稱命名的 Google 日曆且略過重複行程。
 ## 環境準備
 	1. Google 帳號
  	2. Apps Script Cheerio 資料庫 (Script ID: 1ReeQ6WO8kKNxoaA_O0XEQ589cIrRvEBA9qcWpNqdOP17i47u6N9M5Xh0) 用以拆解 HTML 超文字標記語言
   	3. Apps Script API 服務 sheet v4 用來將 xls 轉換為 Google 試算表
 ## 演示
 已設定自動觸發器的共用日曆：
 https://calendar.google.com/calendar/u/0?cid=NWViZWUxNjg0YTZkNWYyMjg4Y2QxNDA0MzY4NDczYWRkZjlhYjY5MGQ3MDFlMDg5ZDkyOTQxOGFiN2MxMTZhM0Bncm91cC5jYWxlbmRhci5nb29nbGUuY29t
