using System.IO;
using System.Linq;
using Autodesk.Revit.DB;

namespace YD_RevitTools.LicenseManager.Helpers.AR
{
    public static class SharedParams
    {
        // 應用程式 ID，用於識別由本工具建立的 DirectShape
        public const string AppId = "YD_BIM_Tools";

        public const string GroupName = "YD_BIM_Tools";

        // 核心模板參數
        public const string P_Total = "模板_合計";           // 模板總面積
        public const string P_EffectiveArea = "模板_有效面積";  // 有效模板面積
        public const string P_Category = "模板_類別";         // 模板類別（牆/柱/梁/板）
        public const string P_HostId = "模板_主體ID";        // 宿主元素ID
        public const string P_AnalysisTime = "分析時間";

        static bool _ensuredThisSession = false;

        public static void Ensure(Document doc)
        {
            if (_ensuredThisSession) return;

            var app = doc.Application;
            DefinitionFile defFile = null;
            string old = app.SharedParametersFilename;
            string temp = Path.Combine(Path.GetTempPath(), "YD_BIM_Tools_Params.txt");

            try
            {
                try { if (!string.IsNullOrEmpty(old) && File.Exists(old)) defFile = app.OpenSharedParameterFile(); }
                catch { defFile = null; }
                if (defFile == null)
                {
                    EnsureValidSharedParamFile(temp);
                    app.SharedParametersFilename = temp;
                    defFile = app.OpenSharedParameterFile();
                }
                if (defFile == null) throw new System.InvalidOperationException("無法開啟共用參數檔。");

                var group = defFile.Groups.Cast<DefinitionGroup>().FirstOrDefault(g => g.Name == GroupName)
                            ?? defFile.Groups.Create(GroupName);

                var needs = new (string, ForgeTypeId)[]
                {
                    (P_Total,         SpecTypeId.Area),        // 模板總面積（平方米）
                    (P_EffectiveArea, SpecTypeId.Area),        // 有效模板面積（平方米）
                    (P_Category,      SpecTypeId.String.Text), // 模板類別
                    (P_HostId,        SpecTypeId.String.Text), // 宿主元素ID
                    (P_AnalysisTime,  SpecTypeId.String.Text)  // 分析時間
                };

                void BindAll()
                {
                    var gmCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel);
                    var catSet = app.Create.NewCategorySet(); catSet.Insert(gmCat);
                    var map = doc.ParameterBindings;

                    foreach (var pair in needs)
                    {
                        string name = pair.Item1; ForgeTypeId spec = pair.Item2;
                        var def = FindDefinition(defFile, name)
                                  ?? group.Definitions.Create(new ExternalDefinitionCreationOptions(name, spec));
                        if (!map.Contains(def))
                        {
                            var ib = app.Create.NewInstanceBinding(catSet);
                            map.Insert(def, ib, GroupTypeId.Data);
                        }
                    }
                }

                if (doc.IsModifiable)
                {
                    using (var st = new SubTransaction(doc)) { st.Start(); BindAll(); st.Commit(); }
                }
                else
                {
                    using (var t = new Transaction(doc, "Ensure Shared Parameters")) { t.Start(); BindAll(); t.Commit(); }
                }

                _ensuredThisSession = true;
            }
            finally
            {
                app.SharedParametersFilename = old;
            }
        }

        private static Definition FindDefinition(DefinitionFile file, string name)
        {
            foreach (DefinitionGroup g in file.Groups)
                foreach (Definition d in g.Definitions)
                    if (d.Name == name) return d;
            return null;
        }

        private static void EnsureValidSharedParamFile(string path)
        {
            string header =
                "# This is a Revit shared parameter file." + System.Environment.NewLine +
                "# Do not edit manually." + System.Environment.NewLine +
                "*META VERSION 5 MINVERSION 1" + System.Environment.NewLine;
            if (!File.Exists(path)) { File.WriteAllText(path, header); return; }
            var lines = File.ReadAllLines(path);
            bool ok = lines.Length >= 3 && lines[0].StartsWith("# This is a Revit shared parameter file.")
                     && lines[1].StartsWith("#") && lines[2].StartsWith("*META");
            if (!ok) File.WriteAllText(path, header);
        }
    }
}
