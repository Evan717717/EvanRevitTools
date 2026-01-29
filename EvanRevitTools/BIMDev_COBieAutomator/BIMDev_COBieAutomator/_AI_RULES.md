
# _AI_RULES.md

## 1. 專案目標 (Objective)

**BIMDev_COBieAutomator** 是一個專為 Revit 開發的資料與參數自動化管理工具，主要解決 BIM 模型中 COBie 資料流的雙向維護問題。
根據程式碼分析，其核心功能定義如下：

* 
**Excel 資料注入 (Data Injection)**：讀取特定格式的 Excel (B表)，將 COBie 資料自動寫入對應的 Revit 元件參數中 。


* 
**參數自動化綁定 (Parameter Binding)**：讀取共用參數檔 (.txt)，批次將參數綁定到模型類別 (Category) 或從專案中移除參數 。


* 
**批次處理能力 (Batch Processing)**：具備背景開啟 Revit 檔案、執行注入、存檔並關閉的自動化能力，無需人工逐一開啟模型 。



## 2. 技術堆疊 (Tech Stack)

* 
**語言與框架**：C# / .NET Framework 4.8 (推斷自 `packages.config` 的 `targetFramework="net48"`) 。


* **Revit API 版本**：相容 Revit 2021+ (基於 .NET 4.8 需求)。
* **第三方套件**：
* 
**NPOI (2.5.6)**：用於無須安裝 Office 即可讀取 Excel (.xlsx) 檔案 。


* **System.Memory / System.Buffers**：NPOI 的相依套件，用於優化記憶體操作。


* 
**UI 框架**：WPF (Windows Presentation Foundation) 。



## 3. 核心架構 (Architecture)

### Entry Point (進入點)

* **App.cs (`IExternalApplication`)**：
* 負責應用程式生命週期管理。
* 在 `OnStartup` 中建立名為 "BIM Development" 的 Ribbon Tab 和 "COBie Tools" 面板 。


* 註冊了兩個主要按鈕：
1. 
**資料導入**：對應 `Command` 類別 。


2. 
**共用參數維護**：對應 `CommandParameter` 類別 。






* **Command.cs / CommandParameter.cs (`IExternalCommand`)**：
* 作為 Revit 指令的觸發器。
* 負責取得 `UIApplication` 或 `UIDocument`，並實例化對應的 WPF 視窗 。





### UI 互動與線程模型

* **模式 (Modality)**：
* 採用 **Modal (模態)** 模式。視窗是透過 `ShowDialog()` 呼叫的 。


* **架構意義**：這意味著當視窗開啟時，Revit 的主視窗會被鎖定 (Freeze)，但 API 的執行緒 (Main Thread) 仍然保持活躍並等待視窗關閉。這允許我們直接在 UI 的事件 (如按鈕點擊) 中呼叫 Revit API，而不需要使用 `IExternalEventHandler`。



### Transaction (事務處理)

* 
**屬性設定**：Command 類別標記為 `[Transaction(TransactionMode.Manual)]`，表示手動控制事務 。


* **事務範圍**：
* 事務並非在 Command 開始時建立，而是在 **UI 事件內部** 建立。
* 例如 `MainWindow.xaml.cs` 中的 `BtnRun_Click` 建立了一個名為 "COBie 單機導入" 的事務 。


* 在批次模式下，程式會為每一個背景開啟的文件 (`bgDoc`) 建立獨立的事務 。





## 4. 開發規範 (Guidelines)

基於目前的程式碼快照，未來開發請嚴格遵守以下規範：

1. **UI 與邏輯分離 (Separation of Concerns)**：
* 
*現狀*：目前的業務邏輯 (如 `RunCOBieInjection`, `RunBatchParameterCreation`) 直接寫在 `.xaml.cs` (Code-behind) 中 。


* *規範*：未來應將 Revit API 操作邏輯提取至獨立的 `Service` 或 `Utils` 類別中 (例如 `COBieService.cs`)。UI 僅負責收集使用者輸入與顯示進度，不應包含複雜的 Excel 解析或幾何運算代碼。


2. **Excel 資源釋放 (Resource Management)**：
* 
*現狀*：目前使用了 `using (FileStream ...)` ，這很好。


* *規範*：由於 NPOI 會佔用檔案鎖定，確保在讀取完畢後立即釋放 `IWorkbook` 物件。在批次處理迴圈中，務必確保 Excel 檔案流不會被重複開啟導致 "檔案被另一程序使用" 的錯誤。建議將 Excel 讀取邏輯封裝，讀取完資料即斷開連接，只將資料 (POCO objects) 傳遞給 Revit 處理邏輯。


3. **錯誤處理與日誌 (Error Handling & Logging)**：
* 
*現狀*：多處使用了空的 `try-catch` 或僅用 `MessageBox` 顯示錯誤 。


* *規範*：
* **禁止**使用空的 `catch` 區塊吞噬錯誤 (Swallowing exceptions)，這會增加除錯難度。
* 在批次處理模式下 (`BtnBatchRun_Click`)，除了 `StringBuilder` 的日誌外，建議引入輕量級的 Log 檔案寫入機制 (如寫入 `.txt` 到桌面)，防止程式崩潰時無法查看 UI 上的日誌 。


* 對於 "Vary by Group" 這種進階 API 操作 (`InternalDefinition`)，必須保留詳細的錯誤追蹤，因為這涉及 Revit 內部屬性，容易隨版本變更而失效 。