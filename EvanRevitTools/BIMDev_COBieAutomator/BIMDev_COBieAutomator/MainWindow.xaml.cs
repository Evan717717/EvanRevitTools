using System;
using System.Collections.Generic; // 用來存清單
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using WinForms = System.Windows.Forms;

namespace BIMDev_COBieAutomator
{
    public partial class MainWindow : Window
    {
        private UIDocument _uidoc;
        private Document _doc;

        public MainWindow(UIDocument uidoc)
        {
            InitializeComponent();
            _uidoc = uidoc;
            _doc = uidoc.Document;
            TxtModelName.Text = _doc.Title;
        }

        // 選擇資料夾按鈕
        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            WinForms.FolderBrowserDialog dialog = new WinForms.FolderBrowserDialog();
            dialog.Description = "請選擇包含 B 表 (Excel) 的資料夾";

            if (dialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                string selectedPath = dialog.SelectedPath;
                string[] excelFiles = Directory.GetFiles(selectedPath, "*.xlsx");

                if (excelFiles.Length == 0)
                {
                    MessageBox.Show("該資料夾內找不到任何 Excel (.xlsx) 檔案！");
                    return;
                }

                ListExcelFiles.Items.Clear();
                foreach (string file in excelFiles)
                {
                    CheckBox cb = new CheckBox();
                    // 這裡用 System.IO.Path 避免衝突
                    cb.Content = System.IO.Path.GetFileName(file);
                    cb.Tag = file; // 把完整路徑藏在 Tag 裡
                    cb.IsChecked = true;
                    cb.Margin = new Thickness(2);
                    ListExcelFiles.Items.Add(cb);
                }
                MessageBox.Show($"成功讀取 {excelFiles.Length} 個檔案！");
            }
        }

        // ============================================================
        // ★ 開始執行按鈕的邏輯
        // ============================================================
        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            // 1. 檢查有沒有勾選任何檔案
            List<string> filesToProcess = new List<string>();
            foreach (CheckBox cb in ListExcelFiles.Items)
            {
                if (cb.IsChecked == true)
                {
                    // 從 Tag 拿出完整路徑
                    filesToProcess.Add(cb.Tag.ToString());
                }
            }

            if (filesToProcess.Count == 0)
            {
                MessageBox.Show("請至少勾選一個要處理的 Excel 檔案！");
                return;
            }

            // 2. 準備開始統計
            string report = "【預備導入分析報告】\n";
            int totalElements = 0;

            // 3. 逐一處理每個檔案
            foreach (string filePath in filesToProcess)
            {
                // A. 從檔名分析出 Revit 的類別
                BuiltInCategory targetCategory = GetCategoryFromFilename(filePath);

                // ★ 修正點：使用 BuiltInCategory.INVALID
                if (targetCategory == BuiltInCategory.INVALID)
                {
                    report += $"❌ 無法識別類別: {System.IO.Path.GetFileName(filePath)}\n";
                    continue;
                }

                // B. 去模型裡搜尋該類別的所有元件
                FilteredElementCollector collector = new FilteredElementCollector(_doc);
                IList<Element> elements = collector.OfCategory(targetCategory)
                                                   .WhereElementIsNotElementType()
                                                   .ToElements();

                report += $"✅ {System.IO.Path.GetFileName(filePath)}\n";
                report += $"   對應類別: {targetCategory}\n";
                report += $"   模型中數量: {elements.Count} 個元件\n\n";

                totalElements += elements.Count;
            }

            // 4. 顯示報告
            report += $"----------------------------\n總計找到 {totalElements} 個目標元件。";
            MessageBox.Show(report, "測試連線報告");
        }

        // ============================================================
        // ★ 核心方法：檔名翻譯機
        // ============================================================
        private BuiltInCategory GetCategoryFromFilename(string filePath)
        {
            string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);

            // 抓最後一個橫線後面的字
            // 例如 "設備清單_機電-Pipe Accessories" -> 抓 "Pipe Accessories"
            string categoryName = "";
            if (fileName.Contains("-"))
            {
                categoryName = fileName.Substring(fileName.LastIndexOf('-') + 1).Trim();
            }
            else
            {
                // 如果檔名沒有橫線，暫時無法判斷，回傳無效
                return BuiltInCategory.INVALID;
            }

            switch (categoryName)
            {
                case "Pipe Accessories": return BuiltInCategory.OST_PipeAccessory;
                case "Plumbing Fixtures": return BuiltInCategory.OST_PlumbingFixtures;
                case "Mechanical Equipment": return BuiltInCategory.OST_MechanicalEquipment;
                case "Communication Devices": return BuiltInCategory.OST_CommunicationDevices;
                case "Conduit Fittings": return BuiltInCategory.OST_ConduitFitting;
                case "Data Devices": return BuiltInCategory.OST_DataDevices;
                case "Duct Accessories": return BuiltInCategory.OST_DuctAccessory;
                case "Electrical Equipment": return BuiltInCategory.OST_ElectricalEquipment;
                case "Electrical Fixtures": return BuiltInCategory.OST_ElectricalFixtures;
                case "Fire Alarm Devices": return BuiltInCategory.OST_FireAlarmDevices;
                case "Lighting Devices": return BuiltInCategory.OST_LightingDevices;
                case "Lighting Fixtures": return BuiltInCategory.OST_LightingFixtures;
                case "Pipe Fittings": return BuiltInCategory.OST_PipeFitting;
                case "Sprinklers": return BuiltInCategory.OST_Sprinklers;
                case "Air Terminals": return BuiltInCategory.OST_DuctTerminal;
                case "Telephone Devices": return BuiltInCategory.OST_TelephoneDevices;
                case "Security Devices": return BuiltInCategory.OST_SecurityDevices;
                // 如果有更多，繼續加...

                // ★ 修正點：使用 BuiltInCategory.INVALID
                default: return BuiltInCategory.INVALID;
            }
        }
    }
}