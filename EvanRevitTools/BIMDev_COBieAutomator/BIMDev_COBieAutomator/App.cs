using System;
using System.Reflection; // 用來抓 DLL 路徑
using System.Collections.Generic;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

namespace BIMDev_COBieAutomator
{
    [Transaction(TransactionMode.Manual)]
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            // 1. 建立 Ribbon Tab (分頁)
            string tabName = "BIM Development";
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch
            {
                // 如果分頁已經存在 (例如被其他外掛建立了)，就忽略錯誤繼續往下
            }

            // 2. 建立 Ribbon Panel (面板)
            RibbonPanel panel = null;
            // 嘗試在該分頁下建立面板，如果重名可能會失敗，所以做個檢查
            List<RibbonPanel> panels = application.GetRibbonPanels(tabName);
            foreach (RibbonPanel p in panels)
            {
                if (p.Name == "COBie Tools")
                {
                    panel = p;
                    break;
                }
            }
            // 如果沒找到就新建一個
            if (panel == null)
            {
                panel = application.CreateRibbonPanel(tabName, "COBie Tools");
            }

            // 3. 取得目前這支 DLL 的路徑
            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            // =======================================================
            // ★ 按鈕 1: 資料導入 (對應 Command.cs)
            // =======================================================
            PushButtonData btnData1 = new PushButtonData(
                "BtnImportCOBie",       // 按鈕內部名稱 (唯一)
                "COBie\n資料導入",      // 按鈕顯示名稱 (\n 代表換行)
                assemblyPath,           // DLL 路徑
                "BIMDev_COBieAutomator.Command" // 對應的 C# 類別全名 (Namespace.Class)
            );
            btnData1.AvailabilityClassName = "BIMDev_COBieAutomator.CommandAvailability";
            btnData1.ToolTip = "讀取 Excel B表並寫入 Revit 元件參數";

            // =======================================================
            // ★ 按鈕 2: 參數建立 (對應 CommandParameter.cs)
            // =======================================================
            PushButtonData btnData2 = new PushButtonData(
                "BtnCreateParam",       
                "共用參數\n維護",       // ★ 修改這裡：原本是 "參數\n批次建立"
                assemblyPath,           
                "BIMDev_COBieAutomator.CommandParameter" 
            );
            btnData2.ToolTip = "讀取共用參數檔 (.txt) 並批次綁定、新增或移除模型參數";

            // 4. 將按鈕加入面板
            panel.AddItem(btnData1);
            panel.AddSeparator(); // 加個分隔線
            panel.AddItem(btnData2);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}