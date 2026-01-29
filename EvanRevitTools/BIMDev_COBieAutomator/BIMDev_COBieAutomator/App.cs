using System;
using System.Collections.Generic;
using System.IO; // 為了讀取串流
using System.Reflection; // 為了讀取內嵌資源 & DLL 路徑
using System.Windows.Media.Imaging; // 為了轉換成 Revit 看得懂的圖片格式
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

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
            btnData1.ToolTip = "讀取已進行Mapping之Excel表並寫入Revit元件參數";

            // ★ 設定圖示 (記得圖片 Build Action 要改為 Embedded Resource)
            btnData1.LargeImage = GetEmbeddedImage("Import_32.png");


            // =======================================================
            // ★ 按鈕 2: 參數建立 (對應 CommandParameter.cs)
            // =======================================================
            PushButtonData btnData2 = new PushButtonData(
                "BtnCreateParam",
                "共用參數維護",
                assemblyPath,
                "BIMDev_COBieAutomator.CommandParameter"
            );
            btnData2.ToolTip = "讀取共用參數檔 (.txt) 並批次綁定、新增或移除模型參數";

            // ★ 設定圖示
            btnData2.LargeImage = GetEmbeddedImage("Param_32.png");

            // 4. 將按鈕加入面板
            panel.AddItem(btnData1);
            // panel.AddSeparator(); // 視需求決定要不要分隔線，現在有圖示通常不用分隔線也很好看
            panel.AddItem(btnData2);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        // =======================================================
        // ★ 輔助方法：從 DLL 的嵌入資源中讀取圖片
        // =======================================================
        public BitmapImage GetEmbeddedImage(string name)
        {
            try
            {
                // 1. 取得目前的 Assembly (也就是這個 DLL 自己)
                Assembly assembly = Assembly.GetExecutingAssembly();

                // 2. 組合資源名稱
                // 格式通常是：[專案Namespace].[資料夾名稱].[檔名.副檔名]
                // 請確認你的 Namespace 是 BIMDev_COBieAutomator，且資料夾叫 Resources
                string resourceName = $"BIMDev_COBieAutomator.Resources.{name}";

                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null) return null;

                    BitmapImage image = new BitmapImage();
                    image.BeginInit();
                    image.StreamSource = stream;
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.EndInit();
                    return image;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}