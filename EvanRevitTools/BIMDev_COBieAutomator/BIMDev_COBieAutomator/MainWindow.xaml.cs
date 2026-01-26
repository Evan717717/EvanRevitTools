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
                    // ★ 優化點：過濾掉 Office 的暫存檔 (檔名以 ~$ 開頭)
                    string fileName = System.IO.Path.GetFileName(file);
                    if (fileName.StartsWith("~$")) continue;

                    CheckBox cb = new CheckBox();
                    cb.Content = fileName;
                    cb.Tag = file;
                    cb.IsChecked = true;
                    cb.Margin = new Thickness(2);
                    ListExcelFiles.Items.Add(cb);
                }
                MessageBox.Show($"成功讀取 {ListExcelFiles.Items.Count} 個檔案！"); // 這裡改用 Items.Count 比較準
            }
        }
        // 全選按鈕
        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox cb in ListExcelFiles.Items)
            {
                cb.IsChecked = true;
            }
        }

        // 全不選按鈕
        private void BtnUnselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox cb in ListExcelFiles.Items)
            {
                cb.IsChecked = false;
            }
        }
        // ============================================================
        // ★ 開始執行按鈕的邏輯 (動態欄位版)
        // ============================================================
        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            List<string> filesToProcess = new List<string>();
            foreach (CheckBox cb in ListExcelFiles.Items)
            {
                if (cb.IsChecked == true) filesToProcess.Add(cb.Tag.ToString());
            }

            if (filesToProcess.Count == 0) return;

            // ★ 1. 取得使用者的設定：是否只寫入空白？
            // 如果 RbOnlyFillBlank 被勾選，這個變數就是 true
            bool onlyFillBlank = (RbOnlyFillBlank.IsChecked == true);

            using (Transaction t = new Transaction(_doc, "COBie 自動化導入"))
            {
                t.Start();

                System.Text.StringBuilder log = new System.Text.StringBuilder();
                log.AppendLine("【詳細執行報告】");
                if (onlyFillBlank) log.AppendLine("※ 模式：僅寫入空白參數 (保留既有值)");
                else log.AppendLine("※ 模式：強制覆蓋現有值");
                log.AppendLine("--------------------------------");

                int totalUpdateCount = 0;

                try
                {
                    foreach (string filePath in filesToProcess)
                    {
                        string fileName = System.IO.Path.GetFileName(filePath);
                        BuiltInCategory category = GetCategoryFromFilename(filePath);
                        if (category == BuiltInCategory.INVALID) continue;

                        FilteredElementCollector collector = new FilteredElementCollector(_doc);
                        IList<Element> revitElements = collector.OfCategory(category).WhereElementIsNotElementType().ToElements();

                        if (revitElements.Count == 0)
                        {
                            log.AppendLine($"⚠️ 跳過 (模型無元件): {fileName}");
                            continue;
                        }

                        // 多加一個參數 FileShare.ReadWrite，告訴電腦：「就算別人正在用，也讓我讀一下」
                        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            IWorkbook workbook = new XSSFWorkbook(fs);
                            ISheet sheet = workbook.GetSheetAt(0);

                            IRow headerRow = sheet.GetRow(1);
                            if (headerRow == null) continue;

                            Dictionary<string, int> parameterMap = new Dictionary<string, int>();

                            for (int cellIndex = 0; cellIndex < headerRow.LastCellNum; cellIndex++)
                            {
                                string headerText = GetCellValue(headerRow, cellIndex).Trim();
                                if (!string.IsNullOrEmpty(headerText) && headerText.StartsWith("Cobie_", StringComparison.OrdinalIgnoreCase))
                                {
                                    string suffix = headerText.Substring(6);
                                    string correctRevitName = "COBie_" + suffix;
                                    parameterMap[correctRevitName] = cellIndex;
                                }
                            }

                            if (parameterMap.Count == 0)
                            {
                                log.AppendLine($"❌ 失敗 (無Cobie欄位): {fileName}");
                                continue;
                            }

                            int fileUpdateCount = 0;

                            for (int rowIndex = 3; rowIndex <= sheet.LastRowNum; rowIndex++)
                            {
                                IRow row = sheet.GetRow(rowIndex);
                                if (row == null) continue;

                                string xlsFamily = GetCellValue(row, 0).Trim();
                                string xlsType = GetCellValue(row, 1).Trim();

                                if (string.IsNullOrEmpty(xlsFamily) || string.IsNullOrEmpty(xlsType)) continue;

                                foreach (Element elem in revitElements)
                                {
                                    Element typeElem = _doc.GetElement(elem.GetTypeId());
                                    if (typeElem == null) continue;

                                    string revitType = typeElem.Name;
                                    string revitFamily = typeElem.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString();

                                    if (revitFamily.Equals(xlsFamily, StringComparison.OrdinalIgnoreCase) &&
                                        revitType.Equals(xlsType, StringComparison.OrdinalIgnoreCase))
                                    {
                                        foreach (var map in parameterMap)
                                        {
                                            string paramName = map.Key;
                                            int colIndex = map.Value;
                                            string rawValue = GetCellValue(row, colIndex).Trim();

                                            // ★ 2. 優化邏輯：如果 Excel 是空的，自動變成 "─"
                                            string valueToWrite = string.IsNullOrEmpty(rawValue) ? "─" : rawValue;

                                            Parameter param = elem.LookupParameter(paramName);
                                            if (param == null || param.IsReadOnly) continue;

                                            // ★ 3. 優化邏輯：判斷是否要覆蓋
                                            // 如果模式是「僅寫入空白」，且參數原本就有值 (HasValue 且不是空字串)，就跳過不寫
                                            if (onlyFillBlank)
                                            {
                                                if (param.HasValue && !string.IsNullOrEmpty(param.AsString()))
                                                {
                                                    continue; // 跳過，保留原值
                                                }
                                            }

                                            // 寫入
                                            param.Set(valueToWrite);
                                            fileUpdateCount++;
                                            totalUpdateCount++;
                                        }
                                    }
                                }
                            }
                            log.AppendLine($"✅ 完成: {fileName} (更新 {fileUpdateCount} 筆)");
                        }
                    }

                    t.Commit();
                    log.AppendLine($"\n總計更新了 {totalUpdateCount} 個參數欄位。");
                    MessageBox.Show(log.ToString(), "執行完成");
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    MessageBox.Show($"發生錯誤：\n{ex.ToString()}", "錯誤");
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