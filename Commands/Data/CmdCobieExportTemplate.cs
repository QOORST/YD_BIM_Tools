using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using YD_RevitTools.LicenseManager;
using YD_RevitTools.LicenseManager.Helpers.Data;

namespace YD_RevitTools.LicenseManager.Commands.Data
{
    [Transaction(TransactionMode.ReadOnly)]
    public class CmdCobieExportTemplate : IExternalCommand
    {
        public Result Execute(ExternalCommandData cd, ref string msg, ElementSet set)
        {
            // 檢查授權 - COBie 範本匯出功能
            var licenseManager = YD_RevitTools.LicenseManager.LicenseManager.Instance;
            if (!licenseManager.HasFeatureAccess("COBie.ExportTemplate"))
            {
                TaskDialog.Show("授權限制",
                    "您的授權版本不支援 COBie 範本匯出功能。\n\n" +
                    "請升級至試用版、標準版或專業版以使用此功能。\n\n" +
                    "點擊「授權管理」按鈕以查看或更新授權。");
                return Result.Cancelled;
            }

            var uidoc = cd.Application.ActiveUIDocument;
            if (uidoc == null) { msg = "沒有開啟的文件。"; return Result.Failed; }

            try
            {
                var cfgs = CobieConfigIO.LoadConfig();
                if (cfgs == null || cfgs.Count == 0)
                {
                    TaskDialog.Show("COBie 範本匯出", "尚未建立任何欄位設定，請先於「COBie 欄位管理」新增欄位。");
                    return Result.Cancelled;
                }

                var sfd = new SaveFileDialog { Filter = "CSV (逗號分隔)|*.csv", FileName = $"COBie_Template_{DateTime.Now:yyyyMMdd}.csv" };
                if (sfd.ShowDialog() != DialogResult.OK) return Result.Cancelled;

                var fixedCols = new[] { "UniqueId", "ElementId", "Mark", "FamilyName", "TypeName" }.ToList();
                var importCols = cfgs.Where(c => c.ImportEnabled)
                                     .Select(c => c.DisplayName?.Trim())
                                     .Where(n => !string.IsNullOrWhiteSpace(n))
                                     .Distinct().ToList();
                var headers = fixedCols.Concat(importCols).ToList();

                using (var fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var sw = new StreamWriter(fs, new UTF8Encoding(true)))
                {
                    var tips = new[]
                    {
                        "# ===== COBie 填表說明 =====",
                        "# 1) 表頭請勿更動。請於第2列開始填寫（本檔已提供一列空白示範）。",
                        "# 2) 元件配對優先序：UniqueId > ElementId > Mark（資產編號）。至少填寫其一以提高配對成功率。",
                        "# 3) FamilyName / TypeName / FamilyType 為系統自動產生的人工對照欄位，匯入時一律忽略，請勿填寫或修改。",
                        "# 4) 日期建議格式：YYYY-MM-DD；布林可用 Yes/No、True/False、1/0、是/否。",
                        "# 5) 數值欄位請僅填數字（勿附單位或符號）。留空＝不覆蓋模型原值。",
                        "# 6) 欄位如需新增/調整，請回 Revit「COBie 欄位管理」後重新匯出範本。",
                        "# =================================",
                        ""
                    };
                    foreach (var t in tips) sw.WriteLine(t);

                    sw.WriteLine(Csv(headers));
                    sw.WriteLine(Csv(headers.Select(_ => ""))); // 空白示範列
                }

                TaskDialog.Show("COBie 範本匯出", "已輸出含說明的範本 CSV。");
                return Result.Succeeded;
            }
            catch (Exception ex) { msg = ex.ToString(); return Result.Failed; }
        }

        private static string Csv(IEnumerable<string> cells)
            => string.Join(",", cells.Select(EscapeCsv));
        private static string EscapeCsv(string s)
        {
            if (s == null) return "";
            bool need = s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r");
            s = s.Replace("\"", "\"\"");
            return need ? $"\"{s}\"" : s;
        }
    }
}
