# 鐵道事故事件動態分析看板

## Python 版儀表板（依 Excel 資料）

### 1. 安裝套件

```bash
pip install -r requirements_dashboard.txt
```

或手動安裝：

```bash
pip install pandas openpyxl streamlit plotly
```

### 2. 執行

```bash
streamlit run dashboard.py
```

### 3. 資料來源

- 預設讀取：`C:\Users\hckuo\Desktop\Power BI\tra_test_data.xlsx`
- 可在左側欄修改 Excel 路徑

### 4. Excel 欄位需求

| 欄位名稱     | 說明           |
|-------------|----------------|
| 項次        | 序號           |
| 發生時間    | 事故發生時間   |
| 營運機構    | TRA、THSR 等   |
| 事故事件分類| 重大/一般/異常 |
| 事故事件種類| 事件類型       |
| 原因第一階  | 肇因一階       |
| 原因第二階  | 肇因二階       |
| 原因第三階  | 肇因三階       |

---

## HTML 版（模擬資料）

直接開啟 `index.html` 即可瀏覽。
