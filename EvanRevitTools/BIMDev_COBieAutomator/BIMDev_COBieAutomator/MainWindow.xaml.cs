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

        // 用來記憶上次路徑的變數
        private string _lastExcelPath = "";
        private string _lastRvtPath = "";
        private string _lastSPPath = "";

        public MainWindow(UIApplication uiapp)
        {
            InitializeComponent();

            _uiapp = uiapp;
            _uidoc = uiapp.ActiveUIDocument;

            if (_uidoc != null)
            {
                _doc = _uidoc.Document;
                TxtModelName.Text = _doc.Title;
                TxtModelName.Foreground = Brushes.Blue;
            }
            else
            {
                _doc = null;
                TxtModelName.Text = "(無開啟專案 - 僅限使用批次模式)";
                TxtModelName.Foreground = Brushes.Gray;
                if (BtnRun != null) BtnRun.IsEnabled = false;
            }
        }

        // ============================================================
        // UI 事件區 (左側 Excel) - 強制隔離路徑
        // ============================================================
        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            WinForms.OpenFileDialog dialog = new WinForms.OpenFileDialog();
            dialog.ValidateNames = false;
            dialog.CheckFileExists = false;
            dialog.CheckPathExists = true;
            dialog.FileName = "選擇資料夾";
            dialog.Title = "請選擇包含 B 表 (Excel) 的資料夾";
            dialog.Filter = "資料夾|*.folder";
            dialog.RestoreDirectory = true;

            // ★ 邏輯修正：如果沒有記憶，就強制回桌面，避免被上一個按鈕影響
            if (!string.IsNullOrEmpty(_lastExcelPath) && Directory.Exists(_lastExcelPath))
            {
                dialog.InitialDirectory = _lastExcelPath;
            }
            else
            {
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }

            if (dialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                string selectedPath = Path.GetDirectoryName(dialog.FileName);
                _lastExcelPath = selectedPath; // 寫入記憶

                if (!string.IsNullOrEmpty(selectedPath))
                {
                    string[] excelFiles = Directory.GetFiles(selectedPath, "*.xlsx");
                    ListExcelFiles.Items.Clear();
                    foreach (string file in excelFiles)
                    {
                        if (Path.GetFileName(file).StartsWith("~$")) continue;
                        CheckBox cb = new CheckBox { Content = Path.GetFileName(file), Tag = file, IsChecked = true, Margin = new Thickness(2) };
                        ListExcelFiles.Items.Add(cb);
                    }
                    if (ListExcelFiles.Items.Count == 0) MessageBox.Show("該資料夾內沒有 Excel 檔案");
                }
            }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e) { foreach (CheckBox cb in ListExcelFiles.Items) cb.IsChecked = true; }
        private void BtnUnselectAll_Click(object sender, RoutedEventArgs e) { foreach (CheckBox cb in ListExcelFiles.Items) cb.IsChecked = false; }


        // ============================================================
        // UI 事件區 (批次參數設定) - 強制隔離路徑
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
            dlg.RestoreDirectory = true;

            // ★ 邏輯修正：如果沒有記憶，就強制回桌面
            if (!string.IsNullOrEmpty(_lastSPPath) && Directory.Exists(_lastSPPath))
            {
                dlg.InitialDirectory = _lastSPPath;
            }
            else
            {
                dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }

            if (dlg.ShowDialog() == WinForms.DialogResult.OK)
            {
                _lastSPPath = Path.GetDirectoryName(dlg.FileName); // 寫入記憶

                TxtSharedParamPath_Batch.Text = dlg.FileName;
                TxtSharedParamPath_Batch.Foreground = Brushes.Black;

                try
                {
                    if (_doc != null)
                    {
                        _doc.Application.SharedParametersFilename = dlg.FileName;
                        DefinitionFile spFile = _doc.Application.OpenSharedParameterFile();
                        CmbGroups_Batch.Items.Clear();
                        foreach (DefinitionGroup g in spFile.Groups) CmbGroups_Batch.Items.Add(g.Name);
                    }
                    else
                    {
                        string[] lines = File.ReadAllLines(dlg.FileName);
                        CmbGroups_Batch.Items.Clear();
                        foreach (string line in lines)
                        {
                            if (line.StartsWith("GROUP") && !line.StartsWith("GROUP\tID"))
                            {
                                string[] parts = line.Split('\t');
                                if (parts.Length > 2) CmbGroups_Batch.Items.Add(parts[2]);
                            }
                        }
                    }

                    if (CmbGroups_Batch.Items.Contains("COBie")) CmbGroups_Batch.SelectedItem = "COBie";
                    else if (CmbGroups_Batch.Items.Count > 0) CmbGroups_Batch.SelectedIndex = 0;
                }
                catch { }
            }
        }

        // ============================================================
        // ★ 修改後：選擇 Revit 資料夾 (含數量統計與事件綁定)
        // ============================================================
        private void BtnSelectRvtFolder_Click(object sender, RoutedEventArgs e)
        {
            WinForms.OpenFileDialog dialog = new WinForms.OpenFileDialog();
            dialog.ValidateNames = false;
            dialog.CheckFileExists = false;
            dialog.CheckPathExists = true;
            dialog.FileName = "選擇資料夾";
            dialog.Title = "請選擇包含 Revit 模型 (.rvt) 的資料夾";
            dialog.Filter = "資料夾|*.folder";
            dialog.RestoreDirectory = true;

            // 記憶路徑邏輯
            if (!string.IsNullOrEmpty(_lastRvtPath) && Directory.Exists(_lastRvtPath))
            {
                dialog.InitialDirectory = _lastRvtPath;
            }
            else
            {
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }

            if (dialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                string selectedPath = Path.GetDirectoryName(dialog.FileName);
                _lastRvtPath = selectedPath; // 寫入記憶

                if (!string.IsNullOrEmpty(selectedPath))
                {
                    string[] files = Directory.GetFiles(selectedPath, "*.rvt");
                    ListRvtFiles.Items.Clear();

                    foreach (string f in files)
                    {
                        if (f.Contains(".00") && f.EndsWith(".rvt")) continue;

                        // 建立 CheckBox
                        CheckBox cb = new CheckBox
                        {
                            Content = Path.GetFileName(f),
                            Tag = f,
                            IsChecked = true,
                            Margin = new Thickness(2)
                        };

                        // ★ 關鍵：綁定點擊事件，讓手動勾選也能即時更新數量
                        cb.Click += (s, args) => UpdateRvtCount();

                        ListRvtFiles.Items.Add(cb);
                    }

                    if (ListRvtFiles.Items.Count == 0) MessageBox.Show("該資料夾內沒有 Revit 模型");

                    // ★ 載入完成後，立刻更新一次數量
                    UpdateRvtCount();
                }
            }
        }
        // ============================================================
        // ★ 新增功能：Revit 清單全選/全不選與數量統計
        // ============================================================

        // 1. 更新數量顯示的輔助方法
        private void UpdateRvtCount()
        {
            // 防止介面還沒初始化就執行
            if (ListRvtFiles == null || TxtRvtCount == null) return;

            int total = ListRvtFiles.Items.Count;
            int selected = 0;

            foreach (CheckBox cb in ListRvtFiles.Items)
            {
                if (cb.IsChecked == true) selected++;
            }

            // 更新介面文字
            TxtRvtCount.Text = $"已選取: {selected} / {total}";
        }

        // 2. 全選按鈕事件
        private void BtnSelectAllRvt_Click(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox cb in ListRvtFiles.Items)
            {
                cb.IsChecked = true;
            }
            UpdateRvtCount(); // 更新數字
        }

        // 3. 全不選按鈕事件
        private void BtnUnselectAllRvt_Click(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox cb in ListRvtFiles.Items)
            {
                cb.IsChecked = false;
            }
            UpdateRvtCount(); // 更新數字
        }


        // ============================================================
        // 核心執行區 (以下保持不變)
        // ============================================================

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
            // 1. 基本檢查
            List<string> xlsFiles = GetCheckedFiles(ListExcelFiles);
            List<string> rvtFiles = GetCheckedFiles(ListRvtFiles);
            if (xlsFiles.Count == 0 || rvtFiles.Count == 0) { MessageBox.Show("請確認 Excel 與 Revit 檔案皆已選擇！"); return; }

            // 2. 參數讀取
            bool doCreateParam = CbRunParamCreation.IsChecked == true;
            string spPath = TxtSharedParamPath_Batch.Text;
            string spGroup = CmbGroups_Batch.SelectedItem?.ToString();
            bool spVary = CbVary_Batch.IsChecked == true;
            bool bindToAll = (CmbCategoryMode.SelectedIndex == 1);

            // 3. 防呆檢查
            if (doCreateParam && (string.IsNullOrEmpty(spPath) || File.Exists(spPath) == false))
            {
                MessageBox.Show("請選擇有效的共用參數檔 (.txt)！"); return;
            }

            Autodesk.Revit.ApplicationServices.Application app = _uiapp.Application;
            System.Text.StringBuilder batchLog = new System.Text.StringBuilder();
            batchLog.AppendLine($"【批次全自動作業報告 (同步模式)】 {DateTime.Now}");

            foreach (string rvtPath in rvtFiles)
            {
                string rvtName = Path.GetFileName(rvtPath);
                batchLog.AppendLine($"\n📂 處理模型: {rvtName}");

                // 定義臨時本機檔路徑 (避免直接開中央檔造成鎖定或警告)
                string tempLocalPath = Path.Combine(Path.GetTempPath(), rvtName);

                try
                {
                    // ====================================================
                    // ★ 策略：建立臨時本機檔 -> 修改 -> 同步回中央
                    // ====================================================

                    // 1. 複製檔案到 Temp 資料夾 (這動作等於建立了本機檔)
                    File.Copy(rvtPath, tempLocalPath, true);

                    // 2. 準備開啟選項
                    OpenOptions openOpts = new OpenOptions();
                    // 關閉所有工作集以加速開啟 (Optional)
                    openOpts.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));

                    // ★★★ 修正點：將字串路徑轉換為 ModelPath (解決 CS1503 錯誤) ★★★
                    ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(tempLocalPath);

                    // 使用 ModelPath 開啟模型
                    Document bgDoc = app.OpenDocumentFile(modelPath, openOpts);

                    using (Transaction t = new Transaction(bgDoc, "批次自動化作業"))
                    {
                        t.Start();

                        // --- 步驟 A: 建立參數 ---
                        if (doCreateParam)
                        {
                            string paramLog = RunBatchParameterCreation(bgDoc, app, spPath, spGroup, spVary, bindToAll);
                            batchLog.AppendLine($"   ⚙️ 參數檢查/建立: {paramLog}");
                        }

                        // --- 步驟 B: 導入資料 ---
                        string dataLog = RunCOBieInjection(bgDoc, xlsFiles, RbOnlyFillBlank.IsChecked == true, true);
                        if (dataLog.Contains("嚴重錯誤")) batchLog.AppendLine("   ❌ 資料導入失敗");
                        else batchLog.AppendLine("   ✅ 資料導入完成");

                        t.Commit();
                    }

                    // ====================================================
                    // ★ 關鍵動作：同步回中央 (Synchronize)
                    // ====================================================
                    if (bgDoc.IsWorkshared)
                    {
                        // 設定同步選項
                        SynchronizeWithCentralOptions syncOpts = new SynchronizeWithCentralOptions();
                        syncOpts.Compact = true; // 執行壓縮

                        // 設定要釋放的權限 (全部釋放)
                        RelinquishOptions relinquishOpts = new RelinquishOptions(true);
                        relinquishOpts.CheckedOutElements = true;
                        relinquishOpts.StandardWorksets = true;
                        relinquishOpts.UserWorksets = true;
                        relinquishOpts.FamilyWorksets = true;
                        relinquishOpts.ViewWorksets = true;
                        syncOpts.SetRelinquishOptions(relinquishOpts);

                        // 設定同步時的註解
                        TransactWithCentralOptions transOpts = new TransactWithCentralOptions();

                        // 執行同步
                        bgDoc.SynchronizeWithCentral(transOpts, syncOpts);
                        batchLog.AppendLine("   🔄 已同步回中央檔案 (含壓縮)");
                    }
                    else
                    {
                        // 如果不是工作共用檔，就直接存檔覆蓋回原路徑
                        bgDoc.Close(false);
                        File.Copy(tempLocalPath, rvtPath, true);
                        batchLog.AppendLine("   💾 單機檔已覆蓋儲存");
                    }

                    // 關閉並刪除臨時檔
                    if (bgDoc.IsValidObject) bgDoc.Close(false);
                    if (File.Exists(tempLocalPath)) File.Delete(tempLocalPath);
                }
                catch (Exception ex)
                {
                    batchLog.AppendLine($"   ❌ 嚴重錯誤: {ex.Message}");
                    try { if (File.Exists(tempLocalPath)) File.Delete(tempLocalPath); } catch { }
                }
            }

            MessageBox.Show(batchLog.ToString());
        }

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
                                        string val = GetCellValue(row, kvp.Value);
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

        private string RunBatchParameterCreation(Document doc, Autodesk.Revit.ApplicationServices.Application app, string txtPath, string groupName, bool isVary, bool bindToAll)
        {
            try
            {
                app.SharedParametersFilename = txtPath;
                DefinitionFile spFile = app.OpenSharedParameterFile();
                if (spFile == null) return "無法讀取參數檔";

                DefinitionGroup group = spFile.Groups.get_Item(groupName);
                if (group == null) return $"找不到群組 {groupName}";

                CategorySet catSet = app.Create.NewCategorySet();

                if (bindToAll)
                {
                    foreach (Category cat in doc.Settings.Categories)
                    {
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
                    BuiltInCategory[] cats = new BuiltInCategory[] {
                        BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_MechanicalEquipment,
                        BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_PlumbingFixtures,
                        BuiltInCategory.OST_DuctAccessory, BuiltInCategory.OST_ElectricalEquipment,
                        BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit,
                        BuiltInCategory.OST_Sprinklers, BuiltInCategory.OST_LightingFixtures,
                        BuiltInCategory.OST_CommunicationDevices, BuiltInCategory.OST_FireAlarmDevices,
                        BuiltInCategory.OST_DataDevices, BuiltInCategory.OST_SecurityDevices,
                        BuiltInCategory.OST_Doors, BuiltInCategory.OST_Windows,
                        BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors
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