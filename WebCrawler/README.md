# WebCrawler

## 專案概述
根據客戶後台爬取客戶的銷售資料
根據客戶提供後台帳號密碼與OTP金鑰
先使用selenium獲取登入後cookies等token
在依照此token去使用scrapy發送請求爬取資料並保存
在依照pandas將資料整理成可使用資料並驗證缺漏生成紀錄表方便查看


## 使用套件

- **網頁爬取**：Scrapy, Selenium, selenium-wire, undetected-chromedriver
- **資料處理**：pandas, numpy, openpyxl
- **OTP認證**：pyotp
- **網路**：scrapy-session

## 專案結構

```
WebCrawler/
├── .idea/                      # PyCharm 專案設定
├── chromedriver/               # Chrome WebDriver 檔案
├── image/                      # 文件說明用圖片
├── login_info_temp/            # 臨時登入資訊儲存
├── mySpider/                   # 主要 Scrapy 爬蟲程式碼
├── __pycache__/                # Python 快取檔案
├── log_write.py                # 將爬取下來的檔案整理紀錄表
├── log_write_pnn.py            # 特殊客戶紀錄表
├── main.py                     # 主要入口點
├── main_pnn.py                 # 特殊客戶專用入口點
├── OTP.py                      # 一次性密碼生成
├── Pipfile                     # Pipenv 依賴項
├── Pipfile.lock                # Pipenv 鎖定檔案
├── processes.py                # main.py的核心處理邏輯
├── README.md                   # 專案文件
└── requirements.txt            # Python 依賴項
```

## 功能特點

- **多種站點類型支援**：
  - 單一站點爬取
  - 總站爬取（具分組功能）
  - 多幣別站點爬取

- **認證方式**：
  - 使用者名稱/密碼登入
  - OTP（一次性密碼）認證
  - 金鑰管理

- **資料處理**：
  - 自動資料提取與處理
  - Excel 報表生成
  - 摘要統計與日誌記錄

- **靈活配置**：
  - 日期範圍選擇
  - 站點篩選選項
  - 無頭模式支援



本專案為專有軟體，僅供內部使用。

## 作者

Kenny