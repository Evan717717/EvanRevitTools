

# _AI_RULES.md

## 1. 專案目標 (Objective)

**BIMDev_COBieAutomator** 是一個專為 Revit 開發的資料與參數自動化管理外掛，旨在解決 BIM 模型與 COBie (Construction Operations Building information exchange) 資料流之間的雙向維護問題。

根據程式碼邏輯，其核心功能定義如下：

* 
**Excel 資料注入 (Data Injection)**：透過讀取特定格式的 Excel (B表)，比對族群與類型名稱，將 COBie 資料批次寫入對應的 Revit 元件參數中 。


* 
**參數自動化維護 (Parameter Management)**：讀取共用參數檔 (.txt)，支援批次將參數綁定到指定的模型類別 (Category)，或從專案中移除參數 。


* 
**自動化批次處理 (Batch Automation)**：具備背景開啟多個 Revit 檔案、建立參數、注入資料、並自動判斷執行「同步回中央」或「儲存覆蓋」的能力 。



## 2. 技術堆疊 (Tech Stack)

* 
**Framework**：.NET Framework 4.8 (推斷自 `packages.config` 的 `targetFramework="net48"`) 。


* **Revit API**：相容 Revit 2021+ (基於 .NET 4.8 需求)。
* **UI 框架**：WPF (Windows Presentation Foundation) 使用 XAML 定義介面。
* **第三方套件**：
* 
**NPOI (2.5.6)**：用於無須安裝 Office 即可讀取與解析 Excel (.xlsx) 檔案 。


* 
**System.Memory / System.Buffers**：NPOI 的相依套件，用於優化記憶體操作 。





## 3. 核心架構 (Architecture)

### Entry Point (進入點)

* **App.cs (`IExternalApplication`)**：
* 負責應用程式生命週期管理。
* 在 `OnStartup` 中建立名為 "BIM Development" 的 Ribbon Tab 和 "COBie Tools" 面板 。


* 註冊兩個主要按鈕：
1. 
**資料導入**：對應 `Command.cs` 。


2. 
**共用參數維護**：對應 `CommandParameter.cs` 。






* **Command.cs / CommandParameter.cs (`IExternalCommand`)**：
* 作為 Revit 指令的觸發器。
* 負責取得 `UIApplication` 或 `UIDocument`，並實例化對應的 WPF 視窗 (`MainWindow` 或 `ParameterWindow`) 。





### UI 互動與執行緒模型

* **模式 (Modality)**：
* 採用 **Modal (模態)** 模式。視窗是透過 `ShowDialog()` 呼叫的 。


* **架構意義**：這意味著當視窗開啟時，Revit 的主視窗會被鎖定 (Freeze)，但 API 的執行緒 (Main Thread) 仍然保持活躍。


* **觸發邏輯**：
* 由於視窗處於 Modal 狀態，UI 事件 (如 `BtnRun_Click`) 依然在 Revit 的主執行緒中運行。因此程式碼 **直接呼叫** Revit API (如 `RunCOBieInjection`)，而未使用 `IExternalEventHandler` (External Event) 。





### Transaction (事務處理)

* 
**屬性設定**：Command 類別標記為 `[Transaction(TransactionMode.Manual)]` 。


* **事務範圍**：
* 事務並非在 Command 開始時建立，而是在 **UI 事件內部** 建立。
* 
**單機模式**：在 `BtnRun_Click` 中使用 `using (Transaction t = new Transaction(_doc, "COBie 單機導入"))` 包裹操作 。


* 
**批次模式**：程式會背景開啟文件 (`bgDoc`)，並為該文件建立獨立的事務 `Transaction(bgDoc, "批次自動化作業")` 。


* 
**同步機制**：批次處理中包含進階邏輯，若檔案為工作共用 (Workshared)，則執行 `SynchronizeWithCentral` 並釋放權限；否則執行 `Save` 與 `Close` 。





## 4. 開發規範 (Guidelines)

基於目前的程式碼風格分析，未來協作開發請遵守以下規範：

1. **UI 與邏輯分離 (Separation of Concerns)**：
* *現狀*：目前的業務邏輯 (如 Excel 解析 `RunCOBieInjection`、參數建立 `RunBatchParameterCreation`) 直接寫在 `.xaml.cs` (Code-behind) 中。
* *規範*：未來應將 Revit API 操作邏輯提取至獨立的 `Service` 或 `Manager` 類別中 (例如 `COBieService.cs`, `ParameterManager.cs`)。UI 僅負責收集輸入與顯示進度，不應包含複雜的幾何運算或檔案解析代碼。


2. **錯誤處理 (Error Handling)**：
* 
*現狀*：程式碼中存在多處空的 `catch { }` 區塊 (例如 `GetCellValue` 或迴圈中的 `try-catch`)，這會吞噬錯誤導致除錯困難 。


* *規範*：**禁止**使用空的 Catch 區塊。必須記錄錯誤日誌 (Log) 或在 Debug 模式下拋出例外。特別是在批次處理迴圈中，單一檔案的失敗不應導致整個程序崩潰，但必須將失敗原因記錄在 `StringBuilder` 日誌中。


3. **Excel 資源管理 (Resource Management)**：
* *現狀*：雖然使用了 `using`，但在批次處理大量檔案時，需特別注意 NPOI 的記憶體佔用。
* 
*規範*：讀取 Excel 時，應確保 `FileStream` 設定為 `FileShare.ReadWrite` 以避免檔案鎖定衝突 。讀取完畢後應立即釋放 `IWorkbook` 物件，將資料轉換為 POCO (Plain Old CLR Object) 傳遞給邏輯層，避免長時間持有 Excel 檔案控點。