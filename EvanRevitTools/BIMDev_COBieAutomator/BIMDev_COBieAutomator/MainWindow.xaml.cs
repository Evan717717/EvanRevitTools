using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using WinForms = System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace BIMDev_COBieAutomator
{
    public partial class MainWindow : Window
    {
        private UIApplication _uiapp;
        private UIDocument _uidoc;
        private Document _doc;

        public MainWindow(UIApplication uiapp)
        {
            InitializeComponent();

            _uiapp = uiapp;
            _uidoc = uiapp.ActiveUIDocument;

            // 判斷現在有沒有開圖
            if (_uidoc != null)
            {
                // --- 狀況 A: 有開圖 (正常模式) ---
                _doc = _uidoc.Document;
                TxtModelName.Text = _doc.Title;
                TxtModelName.Foreground = Brushes.Blue;
            }
            else
            {
                // --- 狀況 B: 沒開圖 (Zero Document 模式) ---
                _doc = null;
                TxtModelName.Text = "(無開啟專案 - 僅限使用批次模式)";
                TxtModelName.Foreground = Brushes.Gray;

                // 1. 鎖死「單機執行」按鈕
                if (BtnRun != null) BtnRun.IsEnabled = false;

                // 2. 自動切換到 Tab 2 (批次處理) - 如果你的 XAML 沒命名 TabControl，這行可以省略
                // if (MainTabControl != null) MainTabControl.SelectedIndex = 1; 
            }
        }

        // ============================================================
        // UI 事件區 (左側 Excel)
        // ============================================================
        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            WinForms.FolderBrowserDialog dialog = new WinForms.FolderBrowserDialog();
            if (dialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                string[] excelFiles = Directory.GetFiles(dialog.SelectedPath, "*.xlsx");
                ListExcelFiles.Items.Clear();
                foreach (string file in excelFiles)
                {
                    if (Path.GetFileName(file).StartsWith("~$")) continue;
                    CheckBox cb = new CheckBox { Content = Path.GetFileName(file), Tag = file, IsChecked = true, Margin = new Thickness(2) };
                    ListExcelFiles.Items.Add(cb);
                }
            }
        }
        private void BtnSelectAll_Click(object sender, RoutedEventArgs e) { foreach (CheckBox cb in ListExcelFiles.Items) cb.IsChecked = true; }
        private void BtnUnselectAll_Click(object sender, RoutedEventArgs e) { foreach (CheckBox cb in ListExcelFiles.Items) cb.IsChecked = false; }


        // ============================================================
        // UI 事件區 (批次參數設定)
        // ============================================================
        private void CbRunParamCreation_CheckedChanged(object sender, RoutedEventArgs e)
        {
            bool enable = CbRunParamCreation.IsChecked == true;
            if (GridParamSettings != null) GridParamSettings.IsEnabled = enable;
            if (PanelParamOpts != null) PanelParamOpts.IsEnabled = enable;
        }

        private void BtnSelectSPFile_Click(object sender, RoutedEventArgs e)
        {
            WinForms.OpenFileDialog dlg = new WinForms.OpenFileDialog { Filter = "Shared Parameter (*.txt)|*.txt" };
            if (dlg.ShowDialog() == WinForms.DialogResult.OK)
            {
                TxtSharedParamPath_Batch.Text = dlg.FileName;
                TxtSharedParamPath_Batch.Foreground = Brushes.Black;

                try
                {
                    // 嘗試讀取群組
                    if (_doc != null)
                    {
                        // 如果有開圖，直接用 API 讀最準
                        _doc.Application.SharedParametersFilename = dlg.FileName;
                        DefinitionFile spFile = _doc.Application.OpenSharedParameterFile();
                        CmbGroups_Batch.Items.Clear();
                        foreach (DefinitionGroup g in spFile.Groups) CmbGroups_Batch.Items.Add(g.Name);
                    }
                    else
                    {
                        // ★★★ 修正 Zero Doc 模式下的文字解析 ★★★
                        string[] lines = File.ReadAllLines(dlg.FileName);
                        CmbGroups_Batch.Items.Clear();
                        foreach (string line in lines)
                        {
                            // 格式通常是: GROUP [TAB] ID [TAB] NAME
                            if (line.StartsWith("GROUP") && !line.StartsWith("GROUP\tID"))
                            {
                                string[] parts = line.Split('\t');
                                // parts[0]=GROUP, parts[1]=ID, parts[2]=Name
                                if (parts.Length > 2)
                                {
                                    CmbGroups_Batch.Items.Add(parts[2]); // 抓名字
                                }
                            }
                        }
                    }

                    // 自動選取 COBie
                    if (CmbGroups_Batch.Items.Contains("COBie")) CmbGroups_Batch.SelectedItem = "COBie";
                    else if (CmbGroups_Batch.Items.Count > 0) CmbGroups_Batch.SelectedIndex = 0;
                }
                catch { }
            }
        }

        private void BtnSelectRvtFolder_Click(object sender, RoutedEventArgs e)
        {
            WinForms.FolderBrowserDialog dlg = new WinForms.FolderBrowserDialog();
            if (dlg.ShowDialog() == WinForms.DialogResult.OK)
            {
                string[] files = Directory.GetFiles(dlg.SelectedPath, "*.rvt");
                ListRvtFiles.Items.Clear();
                foreach (string f in files)
                {
                    // 過濾備份檔
                    if (f.Contains(".00") && f.EndsWith(".rvt")) continue;
                    CheckBox cb = new CheckBox { Content = Path.GetFileName(f), Tag = f, IsChecked = true };
                    ListRvtFiles.Items.Add(cb);
                }
            }
        }


        // ============================================================
        // 核心執行區
        // ============================================================

        // 1. 單機執行
        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            if (_doc == null) return;
            List<string> files = GetCheckedFiles(ListExcelFiles);
            if (files.Count == 0) return;

            using (Transaction t = new Transaction(_doc, "COBie 單機導入"))
            {
                t.Start();
                string report = RunCOBieInjection(_doc, files, RbOnlyFillBlank.IsChecked == true, true);
                t.Commit();
                MessageBox.Show(report);
            }
        }

        // 2. 批次執行
        private void BtnBatchRun_Click(object sender, RoutedEventArgs e)
        {
            List<string> xlsFiles = GetCheckedFiles(ListExcelFiles);
            List<string> rvtFiles = GetCheckedFiles(ListRvtFiles);
            if (xlsFiles.Count == 0 || rvtFiles.Count == 0) { MessageBox.Show("請確認 Excel 與 Revit 檔案皆已選擇！"); return; }

            // 參數設定檢查
            bool doCreateParam = CbRunParamCreation.IsChecked == true;
            string spPath = TxtSharedParamPath_Batch.Text;
            string spGroup = CmbGroups_Batch.SelectedItem?.ToString();
            bool spVary = CbVary_Batch.IsChecked == true;
            
            // ★ 新增：讀取綁定模式 (0=MEP, 1=All)
            bool bindToAll = (CmbCategoryMode.SelectedIndex == 1);

            if (doCreateParam && (string.IsNullOrEmpty(spPath) || File.Exists(spPath) == false))
            {
                MessageBox.Show("請選擇有效的共用參數檔 (.txt)！"); return;
            }

            Autodesk.Revit.ApplicationServices.Application app = _uiapp.Application;

            System.Text.StringBuilder batchLog = new System.Text.StringBuilder();
            batchLog.AppendLine($"【批次全自動作業報告】 {DateTime.Now}");

            foreach (string rvtPath in rvtFiles)
            {
                string rvtName = Path.GetFileName(rvtPath);
                batchLog.AppendLine($"\n📂 處理模型: {rvtName}");

                try
                {
                    Document bgDoc = app.OpenDocumentFile(rvtPath);

                    using (Transaction t = new Transaction(bgDoc, "批次自動化作業"))
                    {
                        t.Start();

                        // --- 步驟 A: 建立參數 ---
                        if (doCreateParam)
                        {
                            // ★ 傳入 bindToAll 參數
                            string paramLog = RunBatchParameterCreation(bgDoc, app, spPath, spGroup, spVary, bindToAll);
                            batchLog.AppendLine($"   ⚙️ 參數檢查/建立: {paramLog}");
                        }

                        // --- 步驟 B: 導入資料 ---
                        string dataLog = RunCOBieInjection(bgDoc, xlsFiles, RbOnlyFillBlank.IsChecked == true, true);
                        if (dataLog.Contains("嚴重錯誤")) batchLog.AppendLine("   ❌ 資料導入失敗");
                        else batchLog.AppendLine("   ✅ 資料導入完成");

                        t.Commit();
                    }

                    SaveOptions opts = new SaveOptions { Compact = true };
                    bgDoc.Save(opts);
                    bgDoc.Close(false);
                    batchLog.AppendLine("   💾 存檔並關閉");
                }
                catch (Exception ex)
                {
                    batchLog.AppendLine($"   ❌ 嚴重錯誤: {ex.Message}");
                }
            }

            MessageBox.Show(batchLog.ToString());
        }

        // ============================================================
        // 引擎 1: 資料導入
        // ============================================================
        private string RunCOBieInjection(Document targetDoc, List<string> excelFiles, bool onlyFillBlank, bool ignoreTempFiles)
        {
            System.Text.StringBuilder log = new System.Text.StringBuilder();
            int total = 0;

            try
            {
                foreach (string filePath in excelFiles)
                {
                    string fileName = Path.GetFileName(filePath);
                    if (ignoreTempFiles && fileName.StartsWith("~$")) continue;
                    BuiltInCategory cat = GetCategoryFromFilename(filePath);
                    if (cat == BuiltInCategory.INVALID) continue;

                    FilteredElementCollector col = new FilteredElementCollector(targetDoc);
                    IList<Element> elems = col.OfCategory(cat).WhereElementIsNotElementType().ToElements();
                    if (elems.Count == 0) continue;

                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        IWorkbook workbook = new XSSFWorkbook(fs);
                        ISheet sheet = workbook.GetSheetAt(0);
                        IRow header = sheet.GetRow(1);
                        if (header == null) continue;

                        Dictionary<string, int> map = new Dictionary<string, int>();
                        for (int i = 0; i < header.LastCellNum; i++)
                        {
                            string h = header.GetCell(i)?.ToString().Trim();
                            if (!string.IsNullOrEmpty(h) && h.StartsWith("Cobie_", StringComparison.OrdinalIgnoreCase))
                                map["COBie_" + h.Substring(6)] = i;
                        }
                        if (map.Count == 0) continue;

                        int count = 0;
                        for (int r = 3; r <= sheet.LastRowNum; r++)
                        {
                            IRow row = sheet.GetRow(r);
                            if (row == null) continue;
                            string fName = row.GetCell(0)?.ToString().Trim();
                            string tName = row.GetCell(1)?.ToString().Trim();
                            if (string.IsNullOrEmpty(fName)) continue;

                            foreach (Element el in elems)
                            {
                                Element type = targetDoc.GetElement(el.GetTypeId());
                                if (type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString().Equals(fName, StringComparison.OrdinalIgnoreCase) &&
                                    type.Name.Equals(tName, StringComparison.OrdinalIgnoreCase))
                                {
                                    foreach (var kvp in map)
                                    {
                                        string val = row.GetCell(kvp.Value)?.ToString().Trim();
                                        Parameter p = el.LookupParameter(kvp.Key);
                                        if (p != null && !p.IsReadOnly)
                                        {
                                            if (onlyFillBlank && p.HasValue && !string.IsNullOrEmpty(p.AsString())) continue;
                                            p.Set(string.IsNullOrEmpty(val) ? "─" : val);
                                            count++; total++;
                                        }
                                    }
                                }
                            }
                        }
                        log.AppendLine($"更新 {fileName}: {count} 筆");
                    }
                }
            }
            catch (Exception ex) { log.AppendLine("Error: " + ex.Message); }
            return log.ToString();
        }

        // 輔助方法：讀取儲存格
        private string GetCellValue(IRow row, int cellIndex)
        {
            ICell cell = row.GetCell(cellIndex);
            if (cell == null) return "";

            if (cell.CellType == NPOI.SS.UserModel.CellType.Formula)
            {
                try { return cell.StringCellValue; }
                catch { return cell.NumericCellValue.ToString(); }
            }
            return cell.ToString().Trim();
        }

        // ============================================================
        // ============================================================
        // 引擎 2: 參數建立 (升級版：支援全類別)
        // ============================================================
        private string RunBatchParameterCreation(Document doc, Autodesk.Revit.ApplicationServices.Application app, string txtPath, string groupName, bool isVary, bool bindToAll)
        {
            try
            {
                app.SharedParametersFilename = txtPath;
                DefinitionFile spFile = app.OpenSharedParameterFile();
                if (spFile == null) return "無法讀取參數檔";

                DefinitionGroup group = spFile.Groups.get_Item(groupName);
                if (group == null) return $"找不到群組 {groupName}";

                // 準備 CategorySet
                CategorySet catSet = app.Create.NewCategorySet();

                if (bindToAll)
                {
                    // ★ 模式 A: 綁定所有模型類別 (比照 Phase 3)
                    foreach (Category cat in doc.Settings.Categories)
                    {
                        // 只要允許綁定參數，且是模型類別，就加入
                        // 這裡加一個簡單的過濾，避免系統雜項
                        if (cat.AllowsBoundParameters &&
                            (cat.CategoryType == CategoryType.Model) &&
                            cat.CanAddSubcategory)
                        {
                            catSet.Insert(cat);
                        }
                    }
                }
                else
                {
                    // ★ 模式 B: 僅綁定 MEP 常用類別 (原本的邏輯)
                    BuiltInCategory[] cats = new BuiltInCategory[] {
                        BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_MechanicalEquipment,
                        BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_PlumbingFixtures,
                        BuiltInCategory.OST_DuctAccessory, BuiltInCategory.OST_ElectricalEquipment,
                        BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit,
                        BuiltInCategory.OST_Sprinklers, BuiltInCategory.OST_LightingFixtures,
                        BuiltInCategory.OST_CommunicationDevices, BuiltInCategory.OST_FireAlarmDevices,
                        BuiltInCategory.OST_DataDevices, BuiltInCategory.OST_SecurityDevices,
                        BuiltInCategory.OST_Doors, BuiltInCategory.OST_Windows,
                        BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors // 多加幾個常用的
                    };
                    foreach (var bic in cats)
                    {
                        try { catSet.Insert(doc.Settings.Categories.get_Item(bic)); } catch { }
                    }
                }

                int count = 0;
                foreach (Definition def in group.Definitions)
                {
                    if (doc.ParameterBindings.Contains(def)) continue;

                    InstanceBinding binding = app.Create.NewInstanceBinding(catSet);
                    if (doc.ParameterBindings.Insert(def, binding, BuiltInParameterGroup.PG_DATA))
                    {
                        var map = doc.ParameterBindings;
                        var it = map.ForwardIterator();
                        it.Reset();
                        while (it.MoveNext())
                        {
                            if (it.Key.Name == def.Name)
                            {
                                if (it.Key is InternalDefinition intDef)
                                {
                                    try { intDef.SetAllowVaryBetweenGroups(doc, isVary); } catch { }
                                }
                                break;
                            }
                        }
                        count++;
                    }
                }
                return $"新增 {count} 個參數";
            }
            catch (Exception ex)
            {
                return $"錯誤: {ex.Message}";
            }
        }

        private List<string> GetCheckedFiles(ListBox list)
        {
            List<string> res = new List<string>();
            foreach (CheckBox cb in list.Items)
                if (cb.IsChecked == true) res.Add(cb.Tag.ToString());
            return res;
        }

        private BuiltInCategory GetCategoryFromFilename(string filePath)
        {
            string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string categoryName = "";
            if (fileName.Contains("-")) categoryName = fileName.Substring(fileName.LastIndexOf('-') + 1).Trim();

            switch (categoryName)
            {
                case "Pipe Accessories": return BuiltInCategory.OST_PipeAccessory;
                case "Plumbing Fixtures": return BuiltInCategory.OST_PlumbingFixtures;
                case "Mechanical Equipment": return BuiltInCategory.OST_MechanicalEquipment;
                case "Electrical Fixtures": return BuiltInCategory.OST_ElectricalFixtures;
                case "Electrical Equipment": return BuiltInCategory.OST_ElectricalEquipment;
                case "Duct Accessories": return BuiltInCategory.OST_DuctAccessory;
                case "Communication Devices": return BuiltInCategory.OST_CommunicationDevices;
                case "Conduit Fittings": return BuiltInCategory.OST_ConduitFitting;
                case "Data Devices": return BuiltInCategory.OST_DataDevices;
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