using System;
using System.IO; // 負責讀取檔案路徑
using System.Windows;
using System.Windows.Controls; // 負責 UI 元件 (如 CheckBox)
// 技巧：因為 WPF 和 Forms 都有 "Application" 或 "MessageBox"，為了避免打架，我們幫 Forms 取個外號叫 WinForms
using WinForms = System.Windows.Forms;

namespace BIMDev_COBieAutomator
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        // 當按下「選擇資料夾」按鈕時會執行這裡
        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            // 1. 建立一個「選擇資料夾」的對話框
            WinForms.FolderBrowserDialog dialog = new WinForms.FolderBrowserDialog();
            dialog.Description = "請選擇包含 B 表 (Excel) 的資料夾";

            // 2. 打開對話框，如果使用者按了「確定 (OK)」
            if (dialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                // 3. 取得使用者選的路徑
                string selectedPath = dialog.SelectedPath;

                // 4. 搜尋該資料夾下所有的 .xlsx 檔案
                string[] excelFiles = Directory.GetFiles(selectedPath, "*.xlsx");

                // 5. 如果沒找到檔案，跳出警告
                if (excelFiles.Length == 0)
                {
                    MessageBox.Show("該資料夾內找不到任何 Excel (.xlsx) 檔案！", "找不到檔案", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 6. 清空原本清單上的假資料
                ListExcelFiles.Items.Clear();

                // 7. 將找到的檔案變成 CheckBox 加進去
                foreach (string file in excelFiles)
                {
                    // 建立一個勾選框
                    CheckBox cb = new CheckBox();
                    // 顯示名稱只顯示檔名 (不要顯示長長的路徑)
                    cb.Content = Path.GetFileName(file);
                    // 把「完整路徑」偷偷藏在 Tag 屬性裡，之後程式要讀檔時可以用
                    cb.Tag = file;
                    // 預設全部勾選
                    cb.IsChecked = true;
                    cb.Margin = new Thickness(2); // 讓間距好看一點

                    // 加到介面上
                    ListExcelFiles.Items.Add(cb);
                }

                // 提示使用者
                MessageBox.Show($"成功讀取 {excelFiles.Length} 個檔案！", "讀取完成");
            }
        }
    }
}