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
            // 1. 取得 UI 應用程式
            UIApplication uiapp = commandData.Application;

            try
            {
                // 實例化視窗
                MainWindow myWindow = new MainWindow(uiapp);

                // ★★★ 關鍵修正：改回 ShowDialog() ★★★
                //這會讓視窗保持在 API 的執行環境內，解決 "Starting a transaction..." 的錯誤
                myWindow.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    public class CommandAvailability : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            return true;
        }
    }
}