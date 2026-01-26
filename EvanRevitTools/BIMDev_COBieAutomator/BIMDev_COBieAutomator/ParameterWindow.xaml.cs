using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media; // 用來設定灰色字體
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using WinForms = System.Windows.Forms;

namespace BIMDev_COBieAutomator
{
    public partial class ParameterWindow : Window
    {
        private UIDocument _uidoc;
        private Document _doc;
        private Autodesk.Revit.ApplicationServices.Application _app;

        public ParameterWindow(UIDocument uidoc)
        {
            InitializeComponent();
            _uidoc = uidoc;
            _doc = uidoc.Document;
            _app = _doc.Application;

            LoadAllRevitCategories();
        }

        // 1. 載入 Categories
        private void LoadAllRevitCategories()
        {
            ListCategories.Items.Clear();
            List<Category> allCats = new List<Category>();

            foreach (Category cat in _doc.Settings.Categories)
            {
                if (cat.AllowsBoundParameters)
                {
                    if (!string.IsNullOrEmpty(cat.Name)) allCats.Add(cat);
                }
            }

            var sortedCats = allCats.OrderBy(c => c.Name).ToList();

            foreach (Category cat in sortedCats)
            {
                CheckBox cb = new CheckBox();
                cb.Content = cat.Name;
                cb.Tag = cat;
                if (IsCommonMEPCategory(cat)) cb.IsChecked = true;
                ListCategories.Items.Add(cb);
            }
        }

        private bool IsCommonMEPCategory(Category cat)
        {
            BuiltInCategory bic = (BuiltInCategory)cat.Id.IntegerValue;
            return bic == BuiltInCategory.OST_PipeAccessory ||
                   bic == BuiltInCategory.OST_PipeFitting ||
                   bic == BuiltInCategory.OST_MechanicalEquipment ||
                   bic == BuiltInCategory.OST_ElectricalFixtures ||
                   bic == BuiltInCategory.OST_ElectricalEquipment ||
                   bic == BuiltInCategory.OST_PlumbingFixtures ||
                   bic == BuiltInCategory.OST_DuctAccessory ||
                   bic == BuiltInCategory.OST_DuctTerminal ||
                   bic == BuiltInCategory.OST_CableTray ||
                   bic == BuiltInCategory.OST_Conduit ||
                   bic == BuiltInCategory.OST_Sprinklers ||
                   bic == BuiltInCategory.OST_LightingFixtures;
        }

        // ============================================================
        // ★ 新增：分頁切換與搜尋事件 (作法 B 核心)
        // ============================================================

        // 當使用者切換分頁時觸發
        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 確保是 TabControl 本身觸發的，避免被內部的 ListBox 干擾
            if (e.Source is TabControl)
            {
                // 如果切換到「移除參數」分頁 (TabRemove)，自動掃描專案
                if (TabRemove != null && TabRemove.IsSelected)
                {
                    ScanProjectParameters();
                }
            }
        }

        // 手動點擊「掃描專案」按鈕
        private void BtnScanProject_Click(object sender, RoutedEventArgs e)
        {
            ScanProjectParameters();
        }

        // ★ 核心邏輯：直接掃描專案參數 (不需要 TXT)
        private void ScanProjectParameters()
        {
            ListParamsToRemove.Items.Clear();
            string keyword = TxtRemoveFilter.Text.Trim(); // 取得關鍵字，預設 "COBie"

            // 取得專案中所有的參數綁定
            BindingMap map = _doc.ParameterBindings;
            DefinitionBindingMapIterator it = map.ForwardIterator();
            it.Reset();

            List<Definition> foundDefs = new List<Definition>();

            while (it.MoveNext())
            {
                Definition def = it.Key;

                // 過濾邏輯：
                // 1. 名字包含關鍵字 (忽略大小寫)
                // 2. 如果沒輸入關鍵字，就列出全部 (由使用者自行決定)
                if (!string.IsNullOrEmpty(keyword))
                {
                    if (def.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        foundDefs.Add(def);
                    }
                }
                else
                {
                    foundDefs.Add(def);
                }
            }

            // 排序並顯示
            var sortedDefs = foundDefs.OrderBy(d => d.Name).ToList();

            foreach (Definition def in sortedDefs)
            {
                CheckBox cb = new CheckBox();
                cb.Content = def.Name;
                cb.Tag = def; // 藏 Definition 物件，之後刪除要用
                cb.IsChecked = false; // 預設不勾選
                ListParamsToRemove.Items.Add(cb);
            }
        }

        // ============================================================
        // TXT 檔案讀取區 (Tab 1 專用)
        // ============================================================
        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            WinForms.OpenFileDialog openFileDialog = new WinForms.OpenFileDialog();
            openFileDialog.Filter = "Shared Parameter Files (*.txt)|*.txt";
            if (openFileDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                TxtSharedParamPath.Text = openFileDialog.FileName;
                LoadGroupsFromFile(openFileDialog.FileName);
            }
        }

        private void LoadGroupsFromFile(string path)
        {
            string originalPath = _app.SharedParametersFilename;
            try
            {
                _app.SharedParametersFilename = path;
                DefinitionFile spFile = _app.OpenSharedParameterFile();
                if (spFile != null)
                {
                    CmbGroups.Items.Clear();
                    foreach (DefinitionGroup group in spFile.Groups)
                    {
                        CmbGroups.Items.Add(group.Name);
                    }
                    if (CmbGroups.Items.Contains("COBie")) CmbGroups.SelectedItem = "COBie";
                    else if (CmbGroups.Items.Count > 0) CmbGroups.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"讀取失敗：{ex.Message}");
            }
            finally
            {
                // _app.SharedParametersFilename = originalPath;
            }
        }

        private void CmbGroups_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbGroups.SelectedItem == null) return;
            string groupName = CmbGroups.SelectedItem.ToString();
            RefreshParameterLists(groupName);
        }

        // ★ 修改後的核心邏輯：只負責刷新 Tab 1 (新增清單)
        private void RefreshParameterLists(string groupName)
        {
            ListParamsToAdd.Items.Clear();
            // 注意：這裡不再操作 ListParamsToRemove，因為那是 Tab 2 的工作

            string spPath = TxtSharedParamPath.Text;
            if (string.IsNullOrEmpty(spPath)) return;

            try
            {
                _app.SharedParametersFilename = spPath;
                DefinitionFile spFile = _app.OpenSharedParameterFile();
                if (spFile == null) return;
                DefinitionGroup group = spFile.Groups.get_Item(groupName);
                if (group == null) return;

                foreach (Definition def in group.Definitions)
                {
                    bool existsInProject = _doc.ParameterBindings.Contains(def);

                    // --- 處理「新增清單」 ---
                    CheckBox cbAdd = new CheckBox();
                    cbAdd.Tag = def;

                    if (existsInProject)
                    {
                        // 已存在：變灰、不能勾選
                        cbAdd.Content = $"{def.Name} (已存在)";
                        cbAdd.Foreground = Brushes.Gray;
                        cbAdd.IsEnabled = false;
                        cbAdd.IsChecked = false;
                    }
                    else
                    {
                        // 不存在：正常顯示
                        cbAdd.Content = def.Name;
                        cbAdd.Foreground = Brushes.Black;
                        cbAdd.IsEnabled = true;
                        cbAdd.IsChecked = true;
                    }
                    ListParamsToAdd.Items.Add(cbAdd);
                }
            }
            catch { }
        }

        // ============================================================
        // 執行新增
        // ============================================================
        private void BtnCreate_Click(object sender, RoutedEventArgs e)
        {
            CategorySet catSet = _app.Create.NewCategorySet();
            foreach (CheckBox cb in ListCategories.Items)
            {
                if (cb.IsChecked == true) catSet.Insert(cb.Tag as Category);
            }
            if (catSet.IsEmpty) { MessageBox.Show("請至少勾選一個 Category！"); return; }

            List<Definition> defsToAdd = new List<Definition>();
            foreach (CheckBox cb in ListParamsToAdd.Items)
            {
                if (cb.IsEnabled && cb.IsChecked == true)
                {
                    defsToAdd.Add(cb.Tag as Definition);
                }
            }
            if (defsToAdd.Count == 0) { MessageBox.Show("請選擇要新增的參數！"); return; }

            bool isVary = CbVaryByGroup.IsChecked == true;

            using (Transaction t = new Transaction(_doc, "新增參數"))
            {
                t.Start();
                int count = 0;
                try
                {
                    foreach (Definition def in defsToAdd)
                    {
                        if (_doc.ParameterBindings.Contains(def)) continue;

                        InstanceBinding binding = _app.Create.NewInstanceBinding(catSet);
                        if (_doc.ParameterBindings.Insert(def, binding, BuiltInParameterGroup.PG_DATA))
                        {
                            var map = _doc.ParameterBindings;
                            var it = map.ForwardIterator();
                            it.Reset();
                            while (it.MoveNext())
                            {
                                if (it.Key.Name == def.Name)
                                {
                                    if (it.Key is InternalDefinition intDef)
                                    {
                                        try { intDef.SetAllowVaryBetweenGroups(_doc, isVary); } catch { }
                                    }
                                    break;
                                }
                            }
                            count++;
                        }
                    }
                    t.Commit();
                    MessageBox.Show($"成功新增 {count} 個參數！");

                    // 執行完畢後，重新整理 Tab 1
                    if (CmbGroups.SelectedItem != null)
                        RefreshParameterLists(CmbGroups.SelectedItem.ToString());
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    MessageBox.Show($"錯誤：{ex.Message}");
                }
            }
        }

        // ============================================================
        // 執行移除 (修改版：搭配 ScanProjectParameters)
        // ============================================================
        private void BtnRemove_Click(object sender, RoutedEventArgs e)
        {
            List<Definition> defsToRemove = new List<Definition>();
            foreach (CheckBox cb in ListParamsToRemove.Items)
            {
                if (cb.IsChecked == true) defsToRemove.Add(cb.Tag as Definition);
            }

            if (defsToRemove.Count == 0) { MessageBox.Show("請勾選要移除的參數！"); return; }

            if (MessageBox.Show($"確定要永久刪除這 {defsToRemove.Count} 個參數嗎？\n這將會清除模型中所有元件上的相關資料！",
                "警告", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            using (Transaction t = new Transaction(_doc, "移除參數"))
            {
                t.Start();
                int count = 0;
                try
                {
                    foreach (Definition def in defsToRemove)
                    {
                        try
                        {
                            _doc.ParameterBindings.Remove(def);
                            count++;
                        }
                        catch { }
                    }
                    t.Commit();
                    MessageBox.Show($"成功移除 {count} 個參數！");

                    // 執行完畢後，重新掃描 Tab 2
                    ScanProjectParameters();

                    // 如果 Tab 1 也有選群組，順便刷新一下 (讓灰色變回黑色)
                    if (CmbGroups.SelectedItem != null)
                        RefreshParameterLists(CmbGroups.SelectedItem.ToString());
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    MessageBox.Show($"錯誤：{ex.Message}");
                }
            }
        }

        // ============================================================
        // 全選/全不選按鈕事件
        // ============================================================
        private void BtnSelectAllParams_Click(object sender, RoutedEventArgs e) { SetAllChecks(ListParamsToAdd, true); }
        private void BtnUnselectAllParams_Click(object sender, RoutedEventArgs e) { SetAllChecks(ListParamsToAdd, false); }

        private void BtnSelectAllCats_Click(object sender, RoutedEventArgs e) { SetAllChecks(ListCategories, true); }
        private void BtnUnselectAllCats_Click(object sender, RoutedEventArgs e) { SetAllChecks(ListCategories, false); }

        private void BtnSelectAllRemove_Click(object sender, RoutedEventArgs e) { SetAllChecks(ListParamsToRemove, true); }
        private void BtnUnselectAllRemove_Click(object sender, RoutedEventArgs e) { SetAllChecks(ListParamsToRemove, false); }

        private void SetAllChecks(ListBox list, bool check)
        {
            foreach (CheckBox cb in list.Items)
            {
                if (cb.IsEnabled) cb.IsChecked = check;
            }
        }
    }
}