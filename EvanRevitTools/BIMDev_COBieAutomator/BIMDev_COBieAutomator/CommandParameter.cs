using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace BIMDev_COBieAutomator
{
    [Transaction(TransactionMode.Manual)]
    public class CommandParameter : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // 取得 UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            // 檢查是否開啟了文件
            if (uidoc == null)
            {
                TaskDialog.Show("錯誤", "請先開啟一個 Revit 專案模型！");
                return Result.Cancelled;
            }

            // 開啟我們的新視窗 (ParameterWindow)
            // 注意：我們還沒建立這個視窗，等一下步驟 2 會做
            ParameterWindow paramWin = new ParameterWindow(uidoc);
            paramWin.ShowDialog();

            return Result.Succeeded;
        }
    }
}