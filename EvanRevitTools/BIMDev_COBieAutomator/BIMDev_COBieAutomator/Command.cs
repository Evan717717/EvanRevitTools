using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace BIMDev_COBieAutomator
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // 1. 取得 UI 應用程式與當前文件
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;

            // =======================================================
            // 關鍵修正：把 uidoc 傳進去！
            // =======================================================
            if (uidoc == null)
            {
                TaskDialog.Show("錯誤", "請先開啟一個 Revit 專案模型 (.rvt) 才能使用此工具！");
                return Result.Cancelled; // 回傳取消，結束程式
            }
            // 2. 實例化我們的視窗 (帶著 uidoc 這張門票)
            MainWindow myWindow = new MainWindow(uidoc);

            // 3. 顯示視窗
            myWindow.ShowDialog();

            return Result.Succeeded;
        }
    }
}