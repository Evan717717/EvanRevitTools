using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace BIMDev_COBieAutomator
{
    // 這是告訴 Revit 處理 Transaction 的方式，Manual 是標準做法
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // 1. 取得 UI 應用程式與當前文件 (雖然目前還沒用到，但保留著是好習慣)
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;

            // =======================================================
            // 關鍵修改區：這裡是連接 MainWindow 的地方
            // =======================================================

            // 2. 實例化 (Instantiate) 我們的視窗
            // 這行意思是：依照 MainWindow 的藍圖，造出一個真正的視窗物件
            MainWindow myWindow = new MainWindow();

            // 3. 顯示視窗
            // ShowDialog() 的意思是「模態視窗 (Modal)」
            // 當這個視窗打開時，使用者不能去點後面的 Revit，必須關閉視窗才能繼續操作
            // 這對於這種資料導入工具來說比較安全
            myWindow.ShowDialog();

            // =======================================================

            // 4. 回傳成功狀態 (視窗關閉後才會跑到這一行)
            return Result.Succeeded;
        }
    }
}