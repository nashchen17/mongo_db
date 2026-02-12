# Mongo_DB 專案

本倉庫包含多個 Python 應用專案。

## 專案列表

### 1. Mongo - Flask + MongoDB 應用

一個簡單的 Flask + MongoDB 應用程式，提供 Excel 檔案上傳、資料管理等功能。

📁 **目錄**: `Mongo/`  
📖 **說明文件**: [Mongo/README.md](Mongo/README.md)

**主要功能**:
- Python (Flask) 後端
- 瀏覽器 UI（上傳 Excel、檢視資料、清除資料庫）
- MongoDB 資料儲存
- Docker + docker-compose 部署

### 2. Taiwan Lotto 649 分析工具

台灣大樂透（Taiwan Lotto 649）歷史開獎數據統計分析工具。

📁 **目錄**: `taiwan_lotto649/`  
📖 **說明文件**: [taiwan_lotto649/README.md](taiwan_lotto649/README.md)

**主要功能**:
- 歷史開獎數據收集（支援爬蟲、JSON匯入）
- 號碼出現頻率統計分析
- 熱門號碼分析
- 智能組合生成（5種策略）
- 分析報告匯出

**快速開始**:
```bash
cd taiwan_lotto649
pip install -r requirements.txt
python lotto_analyzer.py
```

**⚠️ 重要聲明**: 此工具僅供統計分析與學術研究使用，不構成任何投注建議。樂透開獎為獨立隨機事件，歷史頻率不代表未來機率。

## 環境需求

- Python 3.7+
- pip (Python 套件管理工具)
- Docker & Docker Compose (Mongo 專案使用)
- MongoDB (Mongo 專案使用)

## 授權

請參考各專案目錄內的說明文件。

## 貢獻

歡迎提交 Issue 或 Pull Request。

---

**提醒**: 本倉庫的所有工具僅供學習和研究使用，請遵守相關法律法規，理性使用。
