// Commands/AdvancedToolCommand.cs (專業版限定功能範例)
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace YD_RevitTools.LicenseManager.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class AdvancedToolCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // 驗證授權
            var validationResult = LicenseManager.Instance.ValidateLicense();

            if (!validationResult.IsValid)
            {
                TaskDialog.Show("授權錯誤",
                    $"無法執行此功能：{validationResult.Message}");
                return Result.Cancelled;
            }

            // 檢查是否為專業版
            if (validationResult.LicenseInfo.LicenseType != LicenseType.Professional)
            {
                TaskDialog td = new TaskDialog("功能限制");
                td.MainInstruction = "此功能僅限專業版使用";
                td.MainContent =
                    $"您目前使用的是 {validationResult.LicenseInfo.GetLicenseTypeName()}，" +
                    "需要升級到專業版才能使用此進階功能。\n\n" +
                    "請聯繫技術支援進行升級。";
                td.CommonButtons = TaskDialogCommonButtons.Ok;
                td.Show();

                return Result.Cancelled;
            }

            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // 執行進階功能...
                TaskDialog.Show("進階功能", "專業版功能執行成功！");

                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}