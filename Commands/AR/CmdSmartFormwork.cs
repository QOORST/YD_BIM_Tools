using System;
using System.Diagnostics;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using YD_RevitTools.LicenseManager;
using YD_RevitTools.LicenseManager.Helpers.AR;

namespace YD_RevitTools.LicenseManager.Commands.AR
{
    [Transaction(TransactionMode.Manual)]
    public class CmdSmartFormwork : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // 授權檢查
                var licenseManager = LicenseManager.Instance; if (!licenseManager.HasFeatureAccess("SmartFormwork"))
                {
                    return Result.Cancelled;
                }

                var doc = commandData.Application.ActiveUIDocument.Document;
                var uidoc = commandData.Application.ActiveUIDocument;

                Debug.WriteLine("=== 啟動智能模板分析 ===");

                // 確保共用參數存在
                SharedParams.Ensure(doc);

                using (var transaction = new Transaction(doc, "智能模板分析"))
                {
                    transaction.Start();

                    // 執行智能模板分析
                    var result = SmartFormworkEngine.AnalyzeAndGenerate(doc, 18.0); // 預設18mm厚度

                    transaction.Commit();

                    // 顯示結果
                    ShowSmartAnalysisResults(result);
                }

                Debug.WriteLine("=== 智能模板分析完成 ===");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"智能模板分析失敗: {ex.Message}";
                Debug.WriteLine(message);
                Debug.WriteLine($"堆疊追蹤: {ex.StackTrace}");
                return Result.Failed;
            }
        }

        private void ShowSmartAnalysisResults(SmartFormworkResult result)
        {
            var dialog = new TaskDialog("🧠 智能模板分析結果")
            {
                MainInstruction = "基於 PourZone 分群的智能模板分析完成",
                MainContent = GenerateMainSummary(result),
                ExpandedContent = GenerateDetailedReport(result),
                CommonButtons = TaskDialogCommonButtons.Ok,
                DefaultButton = TaskDialogResult.Ok,
                FooterText = "💡 智能引擎特色: PourZone分群 + 空間索引 + 類別規則 + 鄰接檢測 + 旗標排除"
            };

            // 添加額外按鈕
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "查看分群詳情", "顯示各澆置區域的詳細分析");
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "顯示技術說明", "查看智能引擎的技術特色");

            var dialogResult = dialog.Show();
            
            HandleDialogResult(dialogResult, result);
        }

        private string GenerateMainSummary(SmartFormworkResult result)
        {
            var lines = new System.Collections.Generic.List<string>();
            
            lines.Add($"🏗️ 分析構件總數: {result.TotalElements} 件");
            lines.Add($"📐 模板總面積: {result.TotalFormworkArea:F2} m²");
            lines.Add($"🎯 生成模板數量: {result.GeneratedFormworkIds.Count} 個");
            lines.Add($"🗂️ 澆置區域群組: {result.GroupResults.Count} 組");
            lines.Add("");
            
            // 效率統計
            var avgAreaPerElement = result.TotalElements > 0 ? result.TotalFormworkArea / result.TotalElements : 0;
            lines.Add($"📊 平均每構件面積: {avgAreaPerElement:F2} m²");
            
            return string.Join(Environment.NewLine, lines);
        }

        private string GenerateDetailedReport(SmartFormworkResult result)
        {
            var lines = new System.Collections.Generic.List<string>();
            
            lines.Add("🧠 智能引擎技術特色:");
            lines.Add("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            lines.Add("✅ PourZone 分群: 依澆置區域分組處理，提高計算精度");
            lines.Add("✅ 空間索引: BoundingBox 快速鄰居查找，提升效能");
            lines.Add("✅ 類別規則: 依梁/板/柱/牆不同需求精確判斷");
            lines.Add("✅ 鄰接檢測: 自動排除接觸面，避免重複計算");
            lines.Add("✅ 旗標排除: 支援 OnGrade/AgainstSoil/Override");
            lines.Add("");
            
            lines.Add("📋 各群組詳細分析:");
            lines.Add("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            foreach (var group in result.GroupResults)
            {
                lines.Add($"🗂️ 群組: {group.Key}");
                lines.Add($"   📊 構件數量: {group.Value.ElementCount} 件");
                lines.Add($"   📐 群組面積: {group.Value.TotalArea:F2} m²");
                lines.Add($"   🎯 生成模板: {group.Value.FormworkIds.Count} 個");
                lines.Add("");
            }
            
            return string.Join(Environment.NewLine, lines);
        }

        private void HandleDialogResult(TaskDialogResult dialogResult, SmartFormworkResult result)
        {
            switch (dialogResult)
            {
                case TaskDialogResult.CommandLink1:
                    ShowGroupDetails(result);
                    break;
                case TaskDialogResult.CommandLink2:
                    ShowTechnicalDetails();
                    break;
            }
        }

        private void ShowGroupDetails(SmartFormworkResult result)
        {
            var lines = new System.Collections.Generic.List<string>();
            
            lines.Add("🗂️ 澆置區域分群詳情:");
            lines.Add("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            foreach (var group in result.GroupResults)
            {
                lines.Add($"📁 群組 ID: {group.Key}");
                lines.Add($"   🏗️ 構件總數: {group.Value.ElementCount} 件");
                lines.Add($"   📐 總面積: {group.Value.TotalArea:F2} m²");
                lines.Add($"   🎯 模板數量: {group.Value.FormworkIds.Count} 個");
                
                // 構件詳情
                var elementDetails = new System.Collections.Generic.List<string>();
                foreach (var elementResult in group.Value.ElementResults)
                {
                    elementDetails.Add($"      元件 {elementResult.Key}: {elementResult.Value.Area:F2} m² " +
                                     $"({elementResult.Value.ProcessedFaces} 面處理, {elementResult.Value.ExcludedFaces} 面排除)");
                }
                
                if (elementDetails.Count <= 5)
                {
                    lines.AddRange(elementDetails);
                }
                else
                {
                    for (int i = 0; i < 3; i++)
                    {
                        lines.Add(elementDetails[i]);
                    }
                    lines.Add($"      ... 還有 {elementDetails.Count - 3} 個構件");
                }
                
                lines.Add("");
            }
            
            var detailDialog = new TaskDialog("澹置區域分群詳情")
            {
                MainInstruction = "各澆置區域分析詳情",
                MainContent = string.Join(Environment.NewLine, lines),
                CommonButtons = TaskDialogCommonButtons.Ok
            };
            
            detailDialog.Show();
        }

        private void ShowTechnicalDetails()
        {
            var content = @"🧠 智能模板引擎技術說明
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📋 處理流程:
1️⃣ 收集 → 多類別結構元素 (梁/板/柱/牆/基礎)
2️⃣ 分群 → 依 AR_PourZone + Phase 參數分組
3️⃣ 索引 → 建立 BoundingBox 空間索引
4️⃣ 篩選 → 類別+面向規則初篩
5️⃣ 檢測 → 鄰接面接觸檢測
6️⃣ 排除 → 旗標條件排除 (OnGrade/AgainstSoil)
7️⃣ 生成 → 建立 DirectShape 模板

🎯 類別規則 (依用戶需求設計):
• 梁: 兩側 + 底面 (頂面與樓板同澆時不算)
• 板: 底面 + 周邊側緣 (頂面作為平台不算)
• 柱: 四周側面 (與其他構件接觸面除外)
• 牆: 兩側面 + 外露頂頭
• 基礎: 外周側面 (與土接觸面可排除)

⚡ 效能優化:
• 空間索引快速鄰居查找
• BoundingBox 預篩選減少幾何計算
• Boolean 操作僅在必要時執行
• 分群處理避免全局計算

🔧 智能特色:
• 自動識別現澆混凝土構件
• 支援複雜幾何扣除邏輯
• 精確面積計算 (ft² → m²)
• 完整參數記錄和追蹤";

            var techDialog = new TaskDialog("智能引擎技術說明")
            {
                MainInstruction = "SmartFormworkEngine 技術架構",
                MainContent = content,
                CommonButtons = TaskDialogCommonButtons.Ok
            };
            
            techDialog.Show();
        }
    }
}