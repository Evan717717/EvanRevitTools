using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms; // 引用 Windows Forms

namespace MyRevitTools // 請確保命名空間與您的專案名稱一致
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CreatePanelTypesCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // 獲取當前的 Revit 文件
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 💡 關鍵點 1：先檢查是否在族群編輯器中
            if (!doc.IsFamilyDocument)
            {
                TaskDialog.Show("錯誤", "此工具必須在族群編輯器中執行！");
                return Result.Cancelled;
            }

            // 💡 關鍵點 2：建立並顯示我們的 UI 視窗
            // 我們將 Revit 的 Document 物件傳遞給視窗，以便它能操作 Revit 資料
            using (PanelCreatorForm form = new PanelCreatorForm(doc))
            {
                form.ShowDialog(); // 以模態方式顯示視窗，Revit 會暫停等待使用者操作
            }

            return Result.Succeeded;
        }
    }
}