# 如何使用真實歷史數據

本文件說明如何準備和使用台灣大樂透的真實歷史開獎數據。

## 方法一：準備 JSON 格式的歷史數據

### 步驟 1: 取得歷史數據

您可以從以下來源取得台灣大樂透歷史開獎資料：

1. **台灣彩券官方網站**
   - 網址: https://www.taiwanlottery.com.tw
   - 路徑: 首頁 → 大樂透 → 開獎號碼查詢

2. **公開資料集**
   - 政府開放資料平台
   - 第三方彩券數據網站

### 步驟 2: 整理成 JSON 格式

將取得的數據整理成以下 JSON 格式：

```json
[
  {
    "draw_number": 1,
    "date": "2024-01-02",
    "numbers": [5, 12, 23, 31, 38, 45],
    "special": 7
  },
  {
    "draw_number": 2,
    "date": "2024-01-05",
    "numbers": [3, 15, 22, 28, 35, 42],
    "special": 11
  },
  {
    "draw_number": 3,
    "date": "2024-01-09",
    "numbers": [8, 14, 19, 27, 33, 48],
    "special": 2
  }
]
```

**欄位說明：**
- `draw_number`: 期數編號（整數）
- `date`: 開獎日期（字串，格式 YYYY-MM-DD）
- `numbers`: 開獎號碼（陣列，包含6個整數，範圍1-49）
- `special`: 特別號（整數，範圍1-49）

### 步驟 3: 儲存 JSON 檔案

將整理好的數據儲存為 `historical_data.json` 檔案，放在與 `lotto_analyzer.py` 相同的目錄下。

### 步驟 4: 修改程式碼載入真實數據

編輯 `lotto_analyzer.py` 中的 `main()` 函數：

```python
def main():
    analyzer = TaiwanLotto649Analyzer()
    
    # 註解掉模擬數據，改用真實數據
    # analyzer.fetch_historical_data(max_draws=500)
    
    # 從檔案載入真實數據
    analyzer.load_from_file('historical_data.json')
    
    # 其他程式碼保持不變...
    analyzer.analyze_frequency()
    analyzer.print_frequency_report()
    # ...
```

### 步驟 5: 執行分析

```bash
python lotto_analyzer.py
```

## 方法二：使用 CSV 轉換工具（需自行實作）

如果您的數據是 CSV 格式，可以撰寫簡單的轉換腳本：

```python
import csv
import json

def csv_to_json(csv_file, json_file):
    """將 CSV 格式的歷史數據轉換為 JSON 格式"""
    data = []
    
    with open(csv_file, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # 假設 CSV 格式為：
            # draw_number,date,num1,num2,num3,num4,num5,num6,special
            entry = {
                'draw_number': int(row['draw_number']),
                'date': row['date'],
                'numbers': [
                    int(row['num1']),
                    int(row['num2']),
                    int(row['num3']),
                    int(row['num4']),
                    int(row['num5']),
                    int(row['num6'])
                ],
                'special': int(row['special'])
            }
            data.append(entry)
    
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"已轉換 {len(data)} 筆資料")

# 使用範例
csv_to_json('lotto_data.csv', 'historical_data.json')
```

## 方法三：網頁爬蟲（進階）

### 使用 BeautifulSoup

修改 `lotto_analyzer.py` 中的 `_fetch_from_official_api()` 方法：

```python
def _fetch_from_official_api(self, max_draws):
    """從台灣彩券官網爬取數據"""
    
    # 這需要根據官網實際結構調整
    base_url = "https://www.taiwanlottery.com.tw/..."
    
    for draw_num in range(1, max_draws + 1):
        # 發送請求
        response = requests.get(f"{base_url}?draw={draw_num}")
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 解析HTML（需根據實際結構調整）
        # 例如：
        # numbers = [int(td.text) for td in soup.select('.number-cell')]
        
        # 加入延遲避免被封鎖
        time.sleep(1)
```

**注意事項：**
1. 網頁爬蟲需要遵守網站的使用條款
2. 建議加入適當的延遲（例如每次請求間隔1-2秒）
3. 某些網站可能需要處理 AJAX 請求或使用 Selenium

### 使用 Selenium（適用於動態網頁）

```python
from selenium import webdriver
from selenium.webdriver.common.by import By

def scrape_with_selenium():
    driver = webdriver.Chrome()
    driver.get("https://www.taiwanlottery.com.tw/...")
    
    # 等待頁面載入
    time.sleep(2)
    
    # 找出號碼元素（需根據實際HTML調整）
    numbers = driver.find_elements(By.CLASS_NAME, "lotto-number")
    
    driver.quit()
```

## 資料驗證

使用真實數據前，建議先驗證數據的正確性：

```python
def validate_data(data):
    """驗證樂透數據格式"""
    for entry in data:
        # 檢查必要欄位
        assert 'numbers' in entry, "缺少 numbers 欄位"
        assert 'special' in entry, "缺少 special 欄位"
        
        # 檢查號碼數量
        assert len(entry['numbers']) == 6, "號碼數量應為6個"
        
        # 檢查號碼範圍
        for num in entry['numbers']:
            assert 1 <= num <= 49, f"號碼 {num} 超出範圍"
        
        assert 1 <= entry['special'] <= 49, "特別號超出範圍"
        
        # 檢查號碼不重複
        assert len(set(entry['numbers'])) == 6, "號碼有重複"
    
    print(f"✓ 數據驗證通過：{len(data)} 期")

# 使用
with open('historical_data.json', 'r') as f:
    data = json.load(f)
    validate_data(data)
```

## 範例數據

這裡提供一個小型範例數據檔案 `example_data.json`：

```json
[
  {
    "draw_number": 1,
    "date": "2024-01-02",
    "numbers": [3, 12, 23, 31, 38, 45],
    "special": 7
  },
  {
    "draw_number": 2,
    "date": "2024-01-05",
    "numbers": [5, 15, 22, 28, 35, 42],
    "special": 11
  },
  {
    "draw_number": 3,
    "date": "2024-01-09",
    "numbers": [8, 14, 19, 27, 33, 48],
    "special": 2
  }
]
```

## 常見問題

### Q: 需要多少期的歷史數據才有參考價值？

建議至少100期以上，最好500期以上。數據越多，統計結果越穩定。

### Q: 特別號需要計入統計嗎？

可以選擇計入或不計入。目前的程式預設會將特別號也計入統計。如果不想計入，可以修改程式碼：

```python
# 在 load_from_file() 或其他載入數據的地方
for draw in self.draw_history:
    self.all_numbers.extend(draw['numbers'])
    # 註解掉下面這行就不會統計特別號
    # self.all_numbers.append(draw['special'])
```

### Q: 如何確保爬蟲不會被封鎖？

1. 加入適當的延遲（每次請求間隔1-2秒）
2. 設定正確的 User-Agent
3. 尊重網站的 robots.txt
4. 避免在短時間內發送大量請求
5. 考慮使用官方 API（如果有提供）

## 相關資源

- [台灣彩券官方網站](https://www.taiwanlottery.com.tw)
- [Python requests 文件](https://docs.python-requests.org/)
- [BeautifulSoup 文件](https://www.crummy.com/software/BeautifulSoup/bs4/doc/)
- [Selenium 文件](https://selenium-python.readthedocs.io/)

---

**提醒：請遵守相關網站的使用條款和法律規定，理性使用爬蟲技術。**
