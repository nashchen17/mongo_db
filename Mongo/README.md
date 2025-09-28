```markdown
# Python + Flask + MongoDB Excel Import Starter

這是一個簡單的範例專案，提供以下功能：
- Python (Flask) 後端
- 瀏覽器 UI（上傳 Excel、檢視資料、一鍵清除資料庫）
- MongoDB 作為資料庫 (pymongo)
- 支援 Excel 匯入（使用 pandas + openpyxl）
- 使用 Docker + docker-compose 一鍵啟動 app 與 mongo
- 提供一鍵清除資料庫的 API（需確認）

## 專案結構
- app/
  - app.py (Flask 應用)
  - templates/index.html (前端頁面)
  - static/main.js (前端 JS)
- Dockerfile
- docker-compose.yml
- requirements.txt
- .env.example
- README.md

## 快速啟動（建議）
1. 複製專案到本地。
2. 建立 .env（可從 .env.example 修改）或使用 docker-compose 裡的預設。
3. 執行：
   ```
   docker-compose up --build
   ```
   服務會暴露在 http://localhost:5000

## 本機開發（不使用 Docker）
1. 建議建立虛擬環境：
   ```
   python -m venv .venv
   source .venv/bin/activate   # Linux / macOS
   .venv\Scripts\activate      # Windows
   ```
2. 安裝依賴：
   ```
   pip install -r requirements.txt
   ```
3. 設定環境變數（例如連到本機或遠端 MongoDB）：
   ```
   export MONGO_URI="mongodb://localhost:27017/"
   export DB_NAME="mydb"
   export COLLECTION_NAME="items"
   ```
4. 啟動：
   ```
   python app/app.py
   ```
   或使用 gunicorn：
   ```
   gunicorn -w 4 -b 0.0.0.0:5000 app.app:app
   ```

## 使用方式（前端）
1. 開啟瀏覽器到 http://localhost:5000
2. 上傳 Excel（.xlsx）：上傳後會將第一個 sheet 的資料一行一筆存到 MongoDB（欄位名稱由 Excel 欄位決定）
3. 點選「List items」可在頁面顯示目前資料（最多顯示 1000 筆）
4. 點選「Clear DB」會要求二次確認（避免誤刪），按下確認後會呼叫 API 清除整個 collection（drop collection）

## API 範例
- GET /api/items
- POST /api/upload  (form-data, field 名稱: file)
- POST /api/clear   (JSON: {"confirm": true})

範例 cURL 上傳：
```
curl -X POST -F "file=@data.xlsx" http://localhost:5000/api/upload
```

範例 cURL 清除：
```
curl -X POST -H "Content-Type: application/json" -d '{"confirm": true}' http://localhost:5000/api/clear
```

## 注意事項與擴充建議
- 目前 upload 未對欄位型別做嚴格驗證；可依需求加入 schema 驗證（如 pydantic、jsonschema）
- 若要多使用者操作或避免任意清除，可為清除 API 加上簡單認證或 admin token
- 若資料量大，建議加上分頁、索引與批次匯入優化
- 若需要將檔案暫存到磁碟再處理，亦可在 app 中加入上傳處理

如果你要，我可以：
- 幫你把資料欄位 schema 定義好並加入驗證
- 加入使用者認證（簡單 token 或 OAuth）
- 將前端改成更完整的 SPA（React / Vue）
```