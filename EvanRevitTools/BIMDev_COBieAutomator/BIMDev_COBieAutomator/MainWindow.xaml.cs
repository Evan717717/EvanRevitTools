using System;
using System.Collections.Generic; // 用來存清單
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using WinForms = System.Windows.Forms;

// ★★★ 記得補上這兩行，不然讀 Excel 會報錯 ★★★
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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
        // ★ 開始執行按鈕的邏輯 (動態欄位版)
        // ============================================================
        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            // 1. 取得勾選的檔案
            List<string> filesToProcess = new List<string>();
            foreach (CheckBox cb in ListExcelFiles.Items)
            {
                if (cb.IsChecked == true) filesToProcess.Add(cb.Tag.ToString());
            }

            if (filesToProcess.Count == 0) return;

            // 2. 開啟 Transaction
            using (Transaction t = new Transaction(_doc, "COBie 自動化導入"))
            {
                t.Start();

                string logReport = "【導入執行報告】\n";
                int totalUpdateCount = 0;

                try
                {
                    foreach (string filePath in filesToProcess)
                    {
                        // A. 檔名判斷類別
                        BuiltInCategory category = GetCategoryFromFilename(filePath);
                        if (category == BuiltInCategory.INVALID) continue;

                        // B. 抓取模型元件
                        FilteredElementCollector collector = new FilteredElementCollector(_doc);
                        IList<Element> revitElements = collector.OfCategory(category).WhereElementIsNotElementType().ToElements();

                        if (revitElements.Count == 0)
                        {
                            logReport += $"⚠️ 跳過 (無元件): {System.IO.Path.GetFileName(filePath)}\n";
                            continue;
                        }

                        // C. 開啟 Excel
                        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                        {
                            IWorkbook workbook = new XSSFWorkbook(fs);
                            ISheet sheet = workbook.GetSheetAt(0);

                            // ===================================================
                            // ★ 第一階段：建立動態地圖 (掃描標題列)
                            // ===================================================

                            // 根據你的截圖，標題在第 2 列 (Row Index = 1)
                            IRow headerRow = sheet.GetRow(1);
                            if (headerRow == null) continue;

                            // 字典：Key=參數名稱(如 Cobie_製造商), Value=欄位索引(如 6)
                            Dictionary<string, int> parameterMap = new Dictionary<string, int>();

                            // 橫向掃描所有格子
                            for (int cellIndex = 0; cellIndex < headerRow.LastCellNum; cellIndex++)
                            {
                                string headerText = GetCellValue(headerRow, cellIndex);

                                // 只要開頭是 "Cobie_"，就加入地圖
                                if (!string.IsNullOrEmpty(headerText) && headerText.StartsWith("Cobie_"))
                                {
                                    parameterMap[headerText] = cellIndex;
                                }
                            }

                            if (parameterMap.Count == 0)
                            {
                                logReport += $"❌ 失敗 (無Cobie欄位): {System.IO.Path.GetFileName(filePath)}\n";
                                continue;
                            }

                            // ===================================================
                            // ★ 第二階段：讀取資料並寫入
                            // ===================================================

                            // 根據你的截圖，資料從第 4 列開始 (Row Index = 3)
                            for (int rowIndex = 3; rowIndex <= sheet.LastRowNum; rowIndex++)
                            {
                                IRow row = sheet.GetRow(rowIndex);
                                if (row == null) continue;

                                // 固定位置：A欄=Family, B欄=Type
                                string xlsFamily = GetCellValue(row, 0);
                                string xlsType = GetCellValue(row, 1);

                                // 如果 Family 或 Type 是空的，代表這行沒資料，跳過
                                if (string.IsNullOrEmpty(xlsFamily) || string.IsNullOrEmpty(xlsType)) continue;

                                // 在 Revit 元件中尋找匹配者
                                foreach (Element elem in revitElements)
                                {
                                    // 取得 Revit 元件的類型資訊
                                    Element typeElem = _doc.GetElement(elem.GetTypeId());
                                    if (typeElem == null) continue;

                                    string revitType = typeElem.Name;
                                    string revitFamily = typeElem.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString();

                                    // 比對名稱 (忽略大小寫)
                                    if (revitFamily.Equals(xlsFamily, StringComparison.OrdinalIgnoreCase) &&
                                        revitType.Equals(xlsType, StringComparison.OrdinalIgnoreCase))
                                    {
                                        // ★ 這裡是最精彩的地方：看地圖辦事
                                        // 遍歷我們剛剛建立的 Map，有多少 Cobie 欄位就寫多少
                                        foreach (var map in parameterMap)
                                        {
                                            string paramName = map.Key; // 例如 "Cobie_製造商"
                                            int colIndex = map.Value;   // 例如 6

                                            // 去那一行抓資料
                                            string valueToWrite = GetCellValue(row, colIndex);

                                            // 寫入 Revit
                                            bool success = SetParameterValue(elem, paramName, valueToWrite);
                                            if (success) totalUpdateCount++;
                                        }
                                    }
                                }
                            }
                        }
                        logReport += $"✅ 完成: {System.IO.Path.GetFileName(filePath)}\n";
                    }

                    t.Commit();
                    logReport += $"\n----------------\n總計更新了 {totalUpdateCount} 個參數欄位。";
                    MessageBox.Show(logReport, "執行完成");
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    MessageBox.Show($"發生錯誤，已取消變更：\n{ex.Message}", "錯誤");
                }
            }
        }

        // 輔助方法：安全地寫入參數
        private bool SetParameterValue(Element elem, string paramName, string value)
        {
            if (string.IsNullOrEmpty(value)) return false;

            // 嘗試找參數
            Parameter param = elem.LookupParameter(paramName);

            // 如果找不到參數，或參數是唯讀的，就寫不進去
            if (param != null && !param.IsReadOnly)
            {
                // 根據參數類型寫入 (大部分 Cobie 都是文字)
                if (param.StorageType == StorageType.String)
                {
                    param.Set(value);
                    return true;
                }
            }
            return false;
        }

        // 輔助方法：讀取儲存格
        private string GetCellValue(IRow row, int cellIndex)
        {
            ICell cell = row.GetCell(cellIndex);
            if (cell == null) return "";

            // ★ 修正點：加上全名 "NPOI.SS.UserModel.CellType" 避免跟 Revit 撞名
            if (cell.CellType == NPOI.SS.UserModel.CellType.Formula)
            {
                try { return cell.StringCellValue; }
                catch { return cell.NumericCellValue.ToString(); }
            }

            return cell.ToString().Trim();
        }

        // ============================================================
        // ★ 核心方法：檔名翻譯機
        // ============================================================
        private BuiltInCategory GetCategoryFromFilename(string filePath)
        {
            string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);

            // 抓最後一個橫線後面的字
            string categoryName = "";
            if (fileName.Contains("-"))
            {
                categoryName = fileName.Substring(fileName.LastIndexOf('-') + 1).Trim();
            }
            else
            {
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

                default: return BuiltInCategory.INVALID;
            }
        }
    }
}