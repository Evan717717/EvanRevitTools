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
            // 1. 取得 UI 應用程式與當前文件
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;

            // 2. 這是我們未來要呼叫 UI 視窗的地方
            // 目前先用簡單的對話框測試環境是否成功
            TaskDialog.Show("BIM Development", "系統連線成功！準備開始載入 COBie 自動化工具。");

            // 3. 回傳成功狀態
            return Result.Succeeded;
        }
    }
}