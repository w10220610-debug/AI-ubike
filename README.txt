臺東 YouBike 智慧調度工具 v2

【功能】
1. 直接讀取 Excel 配置表，不再把場站數字寫死在程式內。
2. 自動辨識 D1、D2、D3。
3. 可依行政區細分顯示。
4. 分開輸出：場站名稱｜2.0 缺／多幾台｜2.0E 缺／多幾台。
5. 可切換夜班、早班、晚班。
6. 可上傳更新後的 Excel，不必每次修改 Python 程式。
7. 可下載分析結果 CSV。

【Windows 安裝】
1. 安裝 Python 3.11 或 3.12，安裝時勾選 Add Python to PATH。
2. 解壓縮此資料夾。
3. 第一次使用：雙擊「第一次安裝.bat」。
4. 安裝完成後：雙擊「啟動網站.bat」。
5. 瀏覽器通常會自動開啟；若沒有，輸入：
   http://localhost:8501

也可以用終端機手動執行：
   python -m pip install -r requirements.txt
   python -m streamlit run app.py

【檔案說明】
- app.py：新版網站主程式。
- 場站配置基底.xlsx：預設 Excel 基底，可替換成更新版本。
- app0712_original.py：你上傳的原始程式備份。
- requirements.txt：需要安裝的免費套件。
- 第一次安裝.bat：Windows 第一次安裝套件用。
- 啟動網站.bat：Windows 日常啟動用。

【注意】
本版先完成 Excel 基底與人工輸入現況。照片自動辨識車數屬於下一個模組，尚未加入本版。
