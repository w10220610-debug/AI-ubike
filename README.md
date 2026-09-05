# AI-ubike V29 正式版

AI 臺東 YouBike 調度系統，保留原有一般分析、智慧調度、手機／電腦介面，並將
YouBike 場站、即時車數、電池與柱號查詢統一搬到 Python Server。

## V29 穩定化重點

- 電池／柱號：30 秒全使用者共享快取、同站請求合併、最大 4 個外部請求、有限重試、5 分鐘舊資料備援。
- 即時車數：手機不再直連 YouBike；Server 依 Excel 實際場站分批查詢並共享結果。
- Excel：只讀可見工作表，自動辨識實際區域，不再把 D1／D2／D3 當成全國固定格式。
- 智慧調度：只為畫面需要的前 10 個候選站預取電池資料。
- 台東備援：`taitung_fallback.json` 只在沒有可用 Excel 場站時使用，不影響外縣市。

## 安裝與啟動

```text
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

OCR 相關重型套件改為選配，需要時另行安裝：

```text
python -m pip install -r requirements-ocr.txt
```

YouBike 查詢錯誤與快取備援紀錄會寫入 `logs/`。
