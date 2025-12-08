// Commands/LicenseManagementCommand.cs
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using YD_RevitTools.LicenseManager.UI;

namespace YD_RevitTools.LicenseManager.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class LicenseManagementCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                LicenseWindow window = new LicenseWindow();
                window.ShowDialog();
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("錯誤", $"開啟授權管理視窗失敗：{ex.Message}");
                return Result.Failed;
            }
        }
    }
}