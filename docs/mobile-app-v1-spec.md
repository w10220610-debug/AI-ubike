# YouBike 個人調度原生 App｜V1.0 開發規格

- 規格日期：2026-10-11
- 狀態：已確認產品需求；尚未建置原生 App 或授權服務
- 專案策略：新 App 獨立開發，現有 Streamlit main / v29-old-ui 不因此變動
- 更新政策：**使用者手動更新，通知不觸發下載或安裝**

## 1. 產品範圍與基本原則

1. 使用 React Native + Expo 建立 iOS / Android 原生 App；不使用 Streamlit WebView 當作正式介面。
2. 每位調度員帳號獨立；本機保存場站配置、優先場站、個人派工、偏好與快取；不與其他同事同步這些資料。
3. 只有管理員核准的帳號可使用；首次帳密登入後，日常以 Face ID／Touch ID／Android 生物辨識解鎖。生物辨識用於本機解鎖，無法取代伺服器授權。
4. 可採用 Supabase Auth 或等效方案管理帳號、裝置數與授權期限；不得將後端管理密鑰放進 App。
5. 本機 SQLite（必要時 SQLCipher）保存個人資料，SecureStore 保存敏感憑證；支援個人加密備份匯出與還原。
6. 調度服務沿用現有 Python 演算法並逐步拆分 API；實時 YouBike 資料須標示更新時間，無網路時不可冒充即時資料。
7. 先提供帳號授權、生物辨識、電池正反查詢、一般分析、智慧調度、優先場站、手動派工；照片辨識與車號反查列第二階段。

## 2. 手動更新政策（V1.0 必做）

### 2.1 需求定義

- **可自動：** 登入後或 App 返回前景時「只查詢版本中繼資料」，通知使用者有新版；可以發送一般推播通知；設定頁有「檢查更新」。
- **禁止自動：** 背景下載 OTA、預先下載 OTA、偷偷安裝、未經確認重新載入程式，及因收到推播就開始下載。
- **唯一下載入口：** 使用者親自按下「立即更新」並完成相容性檢查，才可以呼叫 OTA 下載流程。
- 需要原生二進位檔更新時，「立即更新」僅開啟經驗證的 App Store／TestFlight／Google Play 官方安裝頁；App 不自行下載或安裝 IPA/APK。
- 使用者按下「稍後再說」即可繼續工作；正在操作派工時不打斷、不強制重新載入。
- 已通過首次帳密登入的使用者，更新過程維持本機資料和必要登入狀態；新版本如需重新授權，應顯示原因。
- 本政策指 App 自行管理的更新。App Store／Google Play／TestFlight 的系統自動更新由使用者裝置／商店設定控制，App 無法替使用者關閉；安裝說明需提示關閉該功能以達到全面手動更新。

### 2.2 發布流程

1. 開發者於 GitHub 開發並完成測試。
2. 發布到 staging/preview 渠道，實機驗證同帳號授權、原有派工資料、電池與調度運作。
3. 管理員明確批准正式版本，才將 OTA 發布至 production 渠道，或發布新的原生安裝版本到平台。
4. 發布版本索引（版本、平台、渠道、runtimeVersion、buildNumber、類型、說明、相容性、下載／商店連結與發布時間）。
5. 對已授權使用者發送「新版本可用」通知；推播不夾帶可執行程式或更新檔。
6. 使用者在 App 版本中心選擇「立即更新」或「稍後再說」。
7. 紀錄下載成功／失敗（不包含個人派工或敏感資料），以便協助排錯。

### 2.3 App 版本中心 UI

- 入口：設定 > 版本與更新；首頁有小型非阻斷提示。
- 顯示：目前 App 版本、build 編號、OTA update ID、可更新版本、發布日期、更新說明、BUG 修復內容、預計更新類型。
- 主要操作：「立即更新」、「稍後再說」、「查看詳情」以及設定頁的「檢查更新」。
- 狀態：檢查中、可更新、下載中（百分比若可取得）、下載完成／重新開啟、更新失敗／重試、目前已是最新版本。
- 推播未授權時，開啟 App 仍須能收到版本提醒。
- 不允許版本通知把正在進行的智慧調度、派工輸入或正在儲存的資料立即中斷。

### 2.4 OTA（相容 JavaScript／樣式／圖片資源）

- Expo updates 設定：updates.enabled = true，updates.checkAutomatically = "NEVER"，updates.useEmbeddedUpdate = true。
- App 開啟與返回前景只能呼叫自己的「版本資訊 API」取得輕量發布中繼資料；不得呼叫 fetchUpdateAsync 或在背景預下載 OTA。
- 點下「立即更新」才檢查 runtimeVersion、渠道、OS 平台、已安裝原生功能是否與 OTA 相容。
- 通過後才調用 expo-updates 的 checkForUpdateAsync()，確定有合適更新後才調用 fetchUpdateAsync()。
- 下載完成應先確保本機派工草稿已保存，呈現「更新已備妥，重新啟動套用」；僅在使用者確認時執行 reloadAsync()，或告知下次自行重啟時生效。
- 網路中斷、找不到相容 OTA、下載錯誤：顯示原因、保留現有版本與個人資料，可重試或稍後再說。
- 若已下載新版，於下次 App 啟動由 expo-updates 選擇已快取更新屬正常行為；但首次下載必須由使用者操作發起，不能先下載再等確認。
- 不可用 OTA 發送新原生權限、未內建的原生模組或與已安裝 runtimeVersion 不相容的程式。

### 2.5 原生更新

- 需要新增／升級原生模組、修改 native 設定、修改 runtimeVersion 時使用原生版本。
- App 查得原生更新後只顯示通知；使用者按「立即更新」才開啟相應官方商店／TestFlight 安裝頁。
- 不以 WebView 或不受信任下載網址引導安裝；需辨識平台並確認版本連結為已配置的可信 URL。
- 不支援系統商店頁面時顯示安裝說明；不宣稱 App 可以直接控制商店的安裝流程。
- 試用期及正式分發管道需分開管理；TestFlight 安裝後的「自動更新」需由測試者自行關閉。

### 2.6 版本相容性與授權安全

- 版本識別欄位：平台（ios/android）、環境渠道（preview/production）、nativeVersion、buildNumber、runtimeVersion、OTA updateId。
- 發布資訊 API 回傳：updateKind（none/ota/native）、displayVersion、releaseNotes、bugFixes、publishedAt、minimumSupportedBuild、runtimeVersion、可信商店連結（原生版）。
- appVersion 和 runtimeVersion 分開管理；版本字串相同不代表原生模組一定相容。
- 通知／更新中繼資料不能授予帳號權限；使用完整調度功能仍須通過帳號授權服務。
- 如有重大安全不相容，可拒絕舊客戶端存取 API 並提示必須手動更新，但仍不得在未經同意下下載或安裝。

### 2.7 個人資料與變更管理

- 任何更新前先提交未儲存的編輯到本機 SQLite；重要資料庫遷移前製作可恢復備份。
- 資料庫 migration 須版本化、具備錯誤處理，避免降版後讀取不相容資料。
- 個人調度配置、派工、優先場站、電量門檻與密碼憑證不得因 OTA／原生更新遭清除。
- 更新紀錄分開顯示「更新內容」與「BUG 修復內容」。
- 版本編號採 V1.0.0 起始，BUG 修復調整 patch，累計五項功能更新可升 minor，原生重大不相容調整 major；公開 releaseNotes 與技術 buildNumber 分開。

## 3. 驗收測試（V1.0 發布門檻）

- [ ] 首次登入正常，日後 Face ID／指紋可解鎖，未授權者無法讀取受保護資料。
- [ ] App 啟動、返回前景、接收推播、顯示更新通知都不會預先下載 OTA。
- [ ] 按「稍後再說」後無下載流量、無重啟，所有調度任務可繼續執行。
- [ ] 按「立即更新」後，相容 OTA 才會檢查、下載，並顯示結果。
- [ ] 新版已下載時不在使用者輸入派工中途擅自 reload；套用前取得明確確認。
- [ ] OTA runtime 不相容時不得下載，正確引導到原生更新渠道。
- [ ] iOS 正式版本導向 App Store；測試版本導向 TestFlight；Android 導向 Google Play 正確應用頁。
- [ ] 斷網、查無更新、HTTP 錯誤、下載失敗仍能繼續使用目前可用版本。
- [ ] 更新前後本機配置、派工、優先場站、電量門檻和授權狀態保留。
- [ ] A 裝置更新不影響 B 裝置個人資料與版本選擇。
- [ ] 發布新版需經管理員批准，測試渠道更新不會誤送正式裝置。
- [ ] 實機驗證 iPhone／Android 系統商店的自動更新設定與 App 內手動更新流程互不混淆。

## 4. 非本次實作範圍

- 本文件僅確認 V1.0 開發規格；尚未建立 Expo App、Supabase 專案、發布 API 或任何正式推播。
- 不修改既有 Streamlit Main 和 V29 的即時運行或發布配置。
- 不加入同事間共用派工、跨帳號資料同步或自動安裝。

## 5. 官方技術依據

- Expo Updates 手動檢查／下載：https://docs.expo.dev/eas-update/download-updates/
- Expo updates 設定：https://docs.expo.dev/versions/latest/sdk/updates/
- EAS Update runtime 相容性：https://docs.expo.dev/eas-update/runtime-versions/
- Apple iPhone 商店自動更新設定：https://support.apple.com/en-asia/guide/iphone/iph98709f167/26/ios/26
- TestFlight 自動更新設定：https://testflight.apple.com/
