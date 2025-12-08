using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Threading;
using YD_RevitTools.LicenseManager;
using YD_RevitTools.LicenseManager.Helpers.AR;

// WPF 別名（避免與 Autodesk.Revit.UI.* 名稱衝突）
using WpfWindow = System.Windows.Window;
using WpfGrid = System.Windows.Controls.Grid;
using WpfStackPanel = System.Windows.Controls.StackPanel;
using WpfComboBox = System.Windows.Controls.ComboBox;
using WpfTextBox = System.Windows.Controls.TextBox;
using WpfLabel = System.Windows.Controls.Label;
using WpfButton = System.Windows.Controls.Button;
using WpfThickness = System.Windows.Thickness;
using WpfOrientation = System.Windows.Controls.Orientation;

namespace YD_RevitTools.LicenseManager.Commands.AR
{
    [Transaction(TransactionMode.Manual)]
    public class CmdPickFace : IExternalCommand
    {
        private HashSet<string> _selectedFaces = new HashSet<string>(); // 已選取的面
        private List<ElementId> _previewElements = new List<ElementId>(); // 預覽元素
        private List<FormworkItem> _createdFormworkItems = new List<FormworkItem>(); // 已建立的模板項目
        private Material _currentMaterial;
        private double _currentThickness;

        public Result Execute(ExternalCommandData data, ref string msg, ElementSet set)
        {
            var uiapp = data.Application;
            var uidoc = uiapp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                // 授權檢查
                var licenseManager = LicenseManager.Instance; if (!licenseManager.HasFeatureAccess("FaceFormwork"))
                {
                    return Result.Cancelled;
                }

                SharedParams.Ensure(doc); // 確保共用參數存在

                // 1) 小視窗：材料 + 厚度
                var dlg = new PickFacePalette(doc);
                new System.Windows.Interop.WindowInteropHelper(dlg) { Owner = uiapp.MainWindowHandle };
                var ok = dlg.ShowDialog();
                if (ok != true) return Result.Cancelled;

                _currentMaterial = dlg.SelectedMaterial;
                _currentThickness = dlg.ThicknessMm;

                // 2) 連續點選平面（ESC 結束）
                var filter = new FaceOnHostFilter(allowFloor: true);
                int created = 0;
                double totalAreaM2 = 0; // 總面積統計

                using (var tg = new TransactionGroup(doc, "面生面"))
                {
                    tg.Start();
                    FormworkEngine.BeginRun();

                    // 將 Transaction 移到迴圈外，提升效能並整合為單一操作
                    using (var t = new Transaction(doc, "面生面"))
                    {
                        t.Start();
                        while (true)
                        {
                            Reference r;
                            try
                            {
                                // 提供更詳細的提示訊息
                                var promptMsg = $"點選要生成模板的『面』（ESC 結束）\n" +
                                              $"已建立: {created} 個 | 總面積: {totalAreaM2:F2} m² | 材料: {_currentMaterial?.Name ?? "預設"} | 厚度: {_currentThickness}mm";
                                r = uidoc.Selection.PickObject(ObjectType.Face, filter, promptMsg);
                            }
                            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                            {
                                // 清理預覽元素
                                ClearPreviewElements(doc);
                                break; // 使用者按下 ESC，結束迴圈
                            }

                            var host = doc.GetElement(r.ElementId);
                            var pf = host?.GetGeometryObjectFromReference(r) as PlanarFace;
                            if (pf == null)
                            {
                                TaskDialog.Show("面生面", "僅支援『平面(PlanarFace)』的面。");
                                continue;
                            }

                            // 1. 先創建即時預覽（類似油漆功能）
                            var previewId = CreatePreviewElement(doc, host, pf, _currentThickness);
                            if (previewId != ElementId.InvalidElementId)
                            {
                                // 立即設定預覽元素的材質和顏色
                                var previewElement = doc.GetElement(previewId) as DirectShape;
                                if (previewElement != null)
                                {
                                    SetElementMaterialAndColor(doc, previewElement, _currentMaterial);
                                }
                                
                                _previewElements.Add(previewId);
                                uidoc.RefreshActiveView();
                                Debug.WriteLine($"🎨 創建即時預覽，元素ID: {previewId.Value}");
                            }

                            // 2. 檢查是否已選取過這個面
                            var faceKey = GetFaceKey(host, pf);
                            if (_selectedFaces.Contains(faceKey))
                            {
                                // 提供視覺反饋，顯示該面已選取
                                var overrides = new OverrideGraphicSettings();
                                overrides.SetProjectionLineColor(new Color(255, 100, 100)); // 紅色邊框
                                overrides.SetProjectionLineWeight(3);
                                doc.ActiveView?.SetElementOverrides(host.Id, overrides);
                                
                                // 短暫顯示後恢復
                                uidoc.RefreshActiveView();
                                System.Threading.Thread.Sleep(500);
                                doc.ActiveView?.SetElementOverrides(host.Id, new OverrideGraphicSettings());
                                uidoc.RefreshActiveView();
                                
                                Debug.WriteLine("該面已經選取過，跳過");
                                continue;
                            }

                            // 只生成用戶點選的面的模板（不使用傳統架構）
                            try
                            {
                                ElementId id = ElementId.InvalidElementId;

                                // 直接為選中的面生成模板
                                id = CreateSingleFaceFormworkDirectly(doc, host, pf, _currentThickness, _currentMaterial);
                                
                                if (id != ElementId.InvalidElementId)
                                {
                                    Debug.WriteLine($"成功為選中面生成模板 ID: {id.Value}");
                                }
                                else
                                {
                                    Debug.WriteLine("單面模板生成失敗");
                                }

                                if (id != ElementId.InvalidElementId) 
                                {
                                    created++;
                                    _selectedFaces.Add(GetFaceKey(host, pf)); // 記錄已選取的面
                                    
                                    // 計算並設定模板面積
                                    var element = doc.GetElement(id);
                                    double areaM2 = 0;
                                    if (element is DirectShape ds)
                                    {
                                        Debug.WriteLine($"🔍 開始處理DirectShape元素 ID: {id.Value}");
                                        
                                        // 先計算面積（包含詳細除錯）
                                        areaM2 = CalculateFormworkAreaWithDebug(doc, ds, pf);
                                        Debug.WriteLine($"📏 計算得到面積: {areaM2:F6} m²");
                                        
                                        // 設定材質和顏色
                                        SetElementMaterialAndColor(doc, ds, _currentMaterial);
                                        
                                        // 設定面積參數和類別 (傳入宿主元素)
                                        SetFormworkAreaParameter(ds, areaM2, host);
                                        totalAreaM2 += areaM2; // 累計總面積
                                        
                                        Debug.WriteLine($"✅ 元素處理完成，累計總面積: {totalAreaM2:F6} m²");
                                        
                                        // 記錄建立的模板項目供後續統計使用
                                        _createdFormworkItems.Add(new FormworkItem
                                        {
                                            ElementId = id,
                                            MaterialName = _currentMaterial?.Name ?? "預設",
                                            Thickness = _currentThickness,
                                            Area = areaM2,
                                            Notes = $"從 {host.Category?.Name ?? "未知"} 面生成"
                                        });
                                    }
                                    
                                    // 立即重繪以顯示新建立的模板（確保用戶能立即看到結果）
                                    uidoc.RefreshActiveView();
                                    
                                    // 短暫高光顯示新建立的模板
                                    try
                                    {
                                        var overrides = new OverrideGraphicSettings();
                                        overrides.SetProjectionLineColor(new Color(0, 255, 0)); // 綠色邊框
                                        overrides.SetProjectionLineWeight(4);
                                        overrides.SetSurfaceTransparency(0); // 完全不透明以突出顯示
                                        doc.ActiveView?.SetElementOverrides(id, overrides);
                                        uidoc.RefreshActiveView();
                                        
                                        // 在背景執行，避免阻塞主線程
                                        var timer = new System.Windows.Threading.DispatcherTimer
                                        {
                                            Interval = TimeSpan.FromMilliseconds(1500)
                                        };
                                        timer.Tick += (sender, args) =>
                                        {
                                            timer.Stop();
                                            try
                                            {
                                                // 恢復正常顯示（透明度設為 15%）
                                                var normalOverrides = new OverrideGraphicSettings();
                                                if (_currentMaterial?.Color.IsValid == true)
                                                {
                                                    normalOverrides.SetSurfaceForegroundPatternColor(_currentMaterial.Color);
                                                    normalOverrides.SetSurfaceBackgroundPatternColor(_currentMaterial.Color);
                                                    normalOverrides.SetProjectionLineColor(_currentMaterial.Color);
                                                    normalOverrides.SetSurfaceTransparency(15);
                                                }
                                                doc.ActiveView?.SetElementOverrides(id, normalOverrides);
                                                uidoc.RefreshActiveView();
                                            }
                                            catch { /* 忽略清理錯誤 */ }
                                        };
                                        timer.Start();
                                    }
                                    catch { /* 忽略高光顯示錯誤 */ }
                                    
                                    Debug.WriteLine($"✅ 成功生成模板並立即顯示，面積: {areaM2:F2} m² | 總計: {created} 個 | 總面積: {totalAreaM2:F2} m²");
                                }
                                else
                                {
                                    // 提供失敗反饋，但不中斷流程
                                    Debug.WriteLine("該面無法生成模板（面不暴露或被其他結構完全遮擋）");
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"生成模板時發生錯誤: {ex.Message}");
                                // 繼續執行，不中斷用戶操作
                            }
                        }
                        t.Commit();
                    }

                    FormworkEngine.EndRun();
                    tg.Assimilate();
                }

                // ESC 完成任務後，顯示數量計算頁面
                if (created > 0)
                {
                    try
                    {
                        // 建立並顯示數量計算對話框
                        var quantityDialog = new FormworkQuantityDialog(doc, _createdFormworkItems);
                        new System.Windows.Interop.WindowInteropHelper(quantityDialog) { Owner = uiapp.MainWindowHandle };
                        quantityDialog.ShowDialog();
                    }
                    catch (Exception dialogEx)
                    {
                        Debug.WriteLine($"數量計算對話框顯示失敗: {dialogEx.Message}");
                        // 數量對話框失敗時，顯示備用的簡化統計 - 確保數據一致性
                        ShowFallbackSummary(created, totalAreaM2);
                    }
                }
                else
                {
                    TaskDialog.Show("面選模板", "未建立任何模板（可能選錯面或全部被裁切）。");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                return Result.Failed;
            }
        }

        // 只允許牆/柱/梁/板
        private class FaceOnHostFilter : ISelectionFilter
        {
            private readonly bool _allowFloor;
            public FaceOnHostFilter(bool allowFloor) { _allowFloor = allowFloor; }

            public bool AllowElement(Element e)
            {
                if (e?.Category?.Id == null) return false;
                long v = e.Category.Id.Value;
                if (v == (long)BuiltInCategory.OST_Walls) return true;
                if (v == (long)BuiltInCategory.OST_StructuralColumns) return true;
                if (v == (long)BuiltInCategory.OST_StructuralFraming) return true;
                if (_allowFloor && v == (long)BuiltInCategory.OST_Floors) return true;
                return false;
            }
            public bool AllowReference(Reference r, XYZ p) => true;
        }

        // 材料 + 厚度的小視窗（中文 UI）
        private class PickFacePalette : WpfWindow
        {
            private readonly Document _doc;
            private readonly WpfComboBox _cmb;
            private readonly WpfTextBox _tbThk;

            public Material SelectedMaterial { get; private set; }
            public double ThicknessMm { get; private set; } = 20.0;

            public PickFacePalette(Document doc)
            {
                _doc = doc;
                Title = "面生面（像油漆）";
                Width = 380; Height = 160;
                WindowStyle = System.Windows.WindowStyle.ToolWindow;
                ResizeMode = System.Windows.ResizeMode.NoResize;
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;

                var root = new WpfGrid { Margin = new WpfThickness(10) };
                Content = root;
                root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
                root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
                root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });

                // row1：材料
                var row1 = new WpfStackPanel { Orientation = WpfOrientation.Horizontal, Margin = new WpfThickness(0, 0, 0, 8) };
                row1.Children.Add(new WpfLabel { Content = "材料：", Width = 60, VerticalAlignment = System.Windows.VerticalAlignment.Center });
                _cmb = new WpfComboBox { Width = 280, IsEditable = false };
                _cmb.Items.Add(new MatItem("＜不指定＞", ElementId.InvalidElementId));
                var mats = new FilteredElementCollector(doc).OfClass(typeof(Material)).Cast<Material>().OrderBy(m => m.Name);
                foreach (var m in mats) _cmb.Items.Add(new MatItem(m.Name, m.Id));
                _cmb.SelectedIndex = 0;
                row1.Children.Add(_cmb);
                root.Children.Add(row1);
                System.Windows.Controls.Grid.SetRow(row1, 0);

                // row2：厚度
                var row2 = new WpfStackPanel { Orientation = WpfOrientation.Horizontal, Margin = new WpfThickness(0, 0, 0, 8) };
                row2.Children.Add(new WpfLabel { Content = "厚度 (mm)：", Width = 60, VerticalAlignment = System.Windows.VerticalAlignment.Center });
                _tbThk = new WpfTextBox { Width = 80, Text = "20" };
                row2.Children.Add(_tbThk);
                root.Children.Add(row2);
                System.Windows.Controls.Grid.SetRow(row2, 1);

                // row3：按鈕
                var row3 = new WpfStackPanel { Orientation = WpfOrientation.Horizontal, HorizontalAlignment = System.Windows.HorizontalAlignment.Right };
                var ok = new WpfButton { Content = "開始點選", Width = 100, Margin = new WpfThickness(0, 0, 8, 0), IsDefault = true };
                var cancel = new WpfButton { Content = "取消", Width = 80, IsCancel = true };

                ok.Click += (s, e) =>
                {
                    if (!double.TryParse(_tbThk.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var mm) || mm <= 0)
                    {
                        System.Windows.MessageBox.Show("請輸入正確的厚度（mm）。", "面生面");
                        return;
                    }
                    ThicknessMm = mm;

                    var item = _cmb.SelectedItem as MatItem;
                    SelectedMaterial = (item != null && item.Id != ElementId.InvalidElementId)
                        ? _doc.GetElement(item.Id) as Material
                        : null;

                    DialogResult = true; Close();
                };
                cancel.Click += (s, e) => { DialogResult = false; Close(); };

                row3.Children.Add(ok); row3.Children.Add(cancel);
                root.Children.Add(row3);
                System.Windows.Controls.Grid.SetRow(row3, 2);
            }

            private class MatItem
            {
                public string Name; public ElementId Id;
                public MatItem(string n, ElementId id) { Name = n; Id = id; }
                public override string ToString() => Name;
            }
        }

        /// <summary>
        /// 使用改進引擎為單個面生成模板
        /// </summary>
        private List<ElementId> CreateSingleFaceFormworkWithImprovedEngine(Document doc, Element host, PlanarFace face, double thicknessMm)
        {
            try
            {
                // 對於單面選擇，我們需要適配改進引擎的邏輯
                // 但只處理選中的面，而不是整個元素的所有面
                
                var formworkIds = new List<ElementId>();
                
                // 取得所有結構元素用於扣除計算
                var allStructuralElements = GetAllStructuralElementsForPickFace(doc);
                
                // 建立聯合實體
                var unionSolid = CreateUnionSolidForPickFace(allStructuralElements);
                if (unionSolid?.Volume <= 1e-6)
                {
                    Debug.WriteLine("無法建立聯合實體，回退到其他方法");
                    return formworkIds;
                }
                
                // 為選中的面生成模板
                var formworkId = CreateFormworkForSingleFace(doc, host, face, unionSolid, thicknessMm);
                if (formworkId != ElementId.InvalidElementId)
                {
                    formworkIds.Add(formworkId);
                }
                
                return formworkIds;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"改進引擎單面模板生成失敗: {ex.Message}");
                return new List<ElementId>();
            }
        }

        private List<Element> GetAllStructuralElementsForPickFace(Document doc)
        {
            // 🚀 重構: 使用統一的 StructuralElementCollector 工具類別
            return StructuralElementCollector.CollectAll(doc, includeFoundation: false);
        }

        private Solid CreateUnionSolidForPickFace(List<Element> elements)
        {
            var allSolids = new List<Solid>();
            foreach (var element in elements)
            {
                var solids = GetElementSolidsForPickFace(element);
                allSolids.AddRange(solids);
            }

            if (allSolids.Count == 0) return null;

            try
            {
                Solid unionResult = allSolids.First();
                for (int i = 1; i < allSolids.Count; i++)
                {
                    try
                    {
                        var nextUnion = BooleanOperationsUtils.ExecuteBooleanOperation(
                            unionResult, allSolids[i], BooleanOperationsType.Union);
                        if (nextUnion?.Volume > 1e-6)
                        {
                            unionResult = nextUnion;
                        }
                    }
                    catch { }
                }
                return unionResult;
            }
            catch
            {
                return null;
            }
        }

        private List<Solid> GetElementSolidsForPickFace(Element element)
        {
            // 🚀 重構: 使用統一的 GeometryExtractor 工具類別
            return GeometryExtractor.GetElementSolids(element);
        }

        private ElementId CreateFormworkForSingleFace(Document doc, Element hostElement, PlanarFace surface, Solid unionSolid, double thicknessMm)
        {
            try
            {
                // 檢查面積
                var area = surface.Area * 0.092903; // 轉換為平方米
                if (area < 0.01) return ElementId.InvalidElementId; // 最小面積 1cm²

                // 從面建立基本模板實體
                var thickness = thicknessMm / 304.8; // 轉換為英尺
                var normal = surface.FaceNormal;
                var curveLoops = surface.GetEdgesAsCurveLoops();
                if (curveLoops.Count == 0) return ElementId.InvalidElementId;

                var extrusionVector = normal.Multiply(thickness);
                var formworkSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                    curveLoops, extrusionVector, thickness);

                if (formworkSolid?.Volume <= 1e-6) return ElementId.InvalidElementId;

                // 使用聯合實體進行精確扣除
                var intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                    formworkSolid, unionSolid, BooleanOperationsType.Intersect);

                Solid finalFormwork = formworkSolid;
                if (intersection?.Volume > 1e-6)
                {
                    double intersectionRatio = intersection.Volume / formworkSolid.Volume;
                    if (intersectionRatio > 0.05) // 5% 以上才扣除
                    {
                        var difference = BooleanOperationsUtils.ExecuteBooleanOperation(
                            formworkSolid, unionSolid, BooleanOperationsType.Difference);
                        if (difference?.Volume > 1e-6)
                        {
                            finalFormwork = difference;
                        }
                    }
                }

                // 建立 DirectShape
                var directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                directShape.ApplicationId = "YD_BIM_Formwork";
                directShape.ApplicationDataId = "ImprovedPickFace";
                directShape.SetShape(new GeometryObject[] { finalFormwork });
                directShape.Name = $"面選改進模板_{hostElement.Id}";

                return directShape.Id;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"單面模板生成失敗: {ex.Message}");
                return ElementId.InvalidElementId;
            }
        }

        /// <summary>
        /// 產生面的唯一識別碼，避免重複選取同一個面
        /// </summary>
        private string GetFaceKey(Element hostElement, PlanarFace face)
        {
            try
            {
                // 使用元素ID + 面的幾何特徵作為唯一識別
                var origin = face.Origin;
                var normal = face.FaceNormal;
                var area = face.Area;
                
                return $"{hostElement.Id.Value}_{origin.X:F3}_{origin.Y:F3}_{origin.Z:F3}_{normal.X:F3}_{normal.Y:F3}_{normal.Z:F3}_{area:F6}";
            }
            catch
            {
                // 如果失敗，至少使用元素ID
                return $"{hostElement.Id.Value}_{DateTime.Now.Ticks}";
            }
        }

        /// <summary>
        /// 設定元素的材料和顏色 - 增強3D視圖顯示效果
        /// </summary>
        private void SetElementMaterialAndColor(Document doc, DirectShape element, Material material)
        {
            try
            {
                Debug.WriteLine($"🎨 開始設定材質: {material?.Name ?? "無材質"} 給元素 {element.Id}");

                // 🚀 重構: 使用統一的 VisualEffectsManager 工具類別
                if (material != null)
                {
                    // 1. 設定元素的材質參數
                    var materialParam = element.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM);
                    if (materialParam != null && !materialParam.IsReadOnly)
                    {
                        materialParam.Set(material.Id);
                        Debug.WriteLine($"✅ 成功設定材質參數: {material.Name}");
                    }
                    
                    // 2. 使用 VisualEffectsManager 設定材質和顏色（包含透明度和視覺效果）
                    VisualEffectsManager.SetFormworkMaterialAndColor(doc, element.Id, material, transparency: 60);
                    
                    Debug.WriteLine($"✅ 材質和顏色設定完成: {material.Name}");
                }
                else
                {
                    Debug.WriteLine($"⚠️ 無材質可設定");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ 設定材質和顏色失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 計算模板面積（平方公尺）- 帶詳細除錯資訊
        /// </summary>
        private double CalculateFormworkAreaWithDebug(Document doc, DirectShape formworkElement, PlanarFace originalFace)
        {
            try
            {
                Debug.WriteLine($"🔍 開始詳細面積計算分析...");
                
                // 🚀 重構: 優先使用原始面面積
                if (originalFace != null)
                {
                    double originalAreaM2 = AreaCalculator.CalculateFacePickedFormworkArea(originalFace);
                    Debug.WriteLine($"📐 原始面面積: {originalAreaM2:F6} m²");
                    
                    if (originalAreaM2 > 0)
                    {
                        Debug.WriteLine($"✅ 使用原始面面積作為模板面積");
                        return originalAreaM2;
                    }
                }

                // 🚀 重構: 嘗試從DirectShape幾何計算
                Debug.WriteLine($"� 分析DirectShape幾何...");
                var solids = GeometryExtractor.GetElementSolids(formworkElement);
                
                if (solids.Count > 0)
                {
                    var largestSolid = GeometryExtractor.GetLargestSolid(solids);
                    if (largestSolid != null)
                    {
                        double calculatedAreaM2 = AreaCalculator.CalculateSimpleFormworkArea(largestSolid);
                        Debug.WriteLine($"� 從幾何計算面積: {calculatedAreaM2:F6} m²");
                        
                        if (calculatedAreaM2 > 0)
                        {
                            Debug.WriteLine($"✅ 使用計算得到的幾何面積");
                            return calculatedAreaM2;
                        }
                    }
                }
                
                Debug.WriteLine($"⚠️ 所有方法都無法計算出有效面積");
                return 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ 計算模板面積失敗: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 計算模板面積（平方公尺）- 基於原始面的實際面積
        /// </summary>
        private double CalculateFormworkArea(DirectShape formworkElement)
        {
            try
            {
                // 方法1：嘗試從參數中讀取面積
                var areaParam = formworkElement.LookupParameter("P_EffectiveArea");
                if (areaParam != null && areaParam.HasValue)
                {
                    double paramArea = areaParam.AsDouble();
                    if (paramArea > 0)
                    {
                        // Revit 內部單位轉換為平方公尺
                        double paramAreaM2 = paramArea * 0.092903; // 1 sq ft = 0.092903 sq m
                        Debug.WriteLine($"✅ 從參數讀取面積: {paramArea:F6} sq ft = {paramAreaM2:F6} m²");
                        return paramAreaM2;
                    }
                }

                // 方法2：計算DirectShape的主要面積（通常是最大的面）
                var geometry = formworkElement.get_Geometry(new Options());
                double maxFaceArea = 0;
                double totalVolume = 0;

                foreach (GeometryObject geomObj in geometry)
                {
                    if (geomObj is Solid solid && solid.Volume > 0)
                    {
                        totalVolume += solid.Volume;
                        
                        // 找到最大的面（通常是模板的主要面）
                        foreach (Face face in solid.Faces)
                        {
                            if (face.Area > maxFaceArea)
                            {
                                maxFaceArea = face.Area;
                            }
                        }
                    }
                }

                // 使用最大面積作為模板面積
                double calculatedAreaM2 = maxFaceArea * 0.092903; // 1 sq ft = 0.092903 sq m

                Debug.WriteLine($"✅ 計算最大面積: {maxFaceArea:F6} sq ft = {calculatedAreaM2:F6} m² (體積: {totalVolume:F6} cu ft)");
                return calculatedAreaM2;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ 計算模板面積失敗: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 設定模板面積參數到共用參數欄位
        /// </summary>
        private void SetFormworkAreaParameter(DirectShape element, double areaM2, Element hostElement)
        {
            try
            {
                Debug.WriteLine($"🔧 開始設定面積參數: {areaM2:F6} m² 到元素 {element.Id.Value}");
                
                // 🚀 重構: 使用 AreaCalculator 的單位轉換方法
                double areaInSquareFeet = AreaCalculator.ConvertToSquareFeet(areaM2);
                Debug.WriteLine($"📐 面積轉換: {areaM2:F6} m² = {areaInSquareFeet:F6} sq ft");

                // 方法1：設定共用參數 "模板_有效面積"
                var effectiveAreaParam = element.LookupParameter(SharedParams.P_EffectiveArea);
                if (effectiveAreaParam != null && !effectiveAreaParam.IsReadOnly)
                {
                    effectiveAreaParam.Set(areaInSquareFeet);
                    Debug.WriteLine($"✅ 成功設定共用參數 '{SharedParams.P_EffectiveArea}': {areaM2:F2} m²");
                }
                else
                {
                    Debug.WriteLine($"⚠️ 無法找到或設定共用參數 '{SharedParams.P_EffectiveArea}'");
                }

                // 方法2：設定共用參數 "模板_合計" (也設定相同值)
                var totalParam = element.LookupParameter(SharedParams.P_Total);
                if (totalParam != null && !totalParam.IsReadOnly)
                {
                    totalParam.Set(areaInSquareFeet);
                    Debug.WriteLine($"✅ 成功設定共用參數 '{SharedParams.P_Total}': {areaM2:F2} m²");
                }
                else
                {
                    Debug.WriteLine($"⚠️ 無法找到或設定共用參數 '{SharedParams.P_Total}'");
                }

                // 方法3：設定類別參數 (從宿主元素判斷)
                var categoryParam = element.LookupParameter(SharedParams.P_Category);
                if (categoryParam != null && !categoryParam.IsReadOnly)
                {
                    // 🔧 修正: 使用統一的類別判斷邏輯 (基於宿主元素)
                    string category = GetFormworkCategoryFromHost(hostElement);
                    categoryParam.Set(category);
                    Debug.WriteLine($"✅ 成功設定共用參數 '{SharedParams.P_Category}': {category}");
                }

                // 方法4：設定宿主ID
                var hostIdParam = element.LookupParameter(SharedParams.P_HostId);
                if (hostIdParam != null && !hostIdParam.IsReadOnly)
                {
                    hostIdParam.Set(hostElement.Id.ToString());
                    Debug.WriteLine($"✅ 成功設定共用參數 '{SharedParams.P_HostId}': {hostElement.Id}");
                }

                // 方法5：設定厚度參數
                SetThicknessParameter(element, _currentThickness);

                // 方法6：設定材質參數
                SetMaterialParameter(element, _currentMaterial);

                // 方法7：設定分析時間
                var timeParam = element.LookupParameter(SharedParams.P_AnalysisTime);
                if (timeParam != null && !timeParam.IsReadOnly)
                {
                    var currentTime = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    timeParam.Set(currentTime);
                    Debug.WriteLine($"✅ 成功設定共用參數 '{SharedParams.P_AnalysisTime}': {currentTime}");
                }

                Debug.WriteLine($"🎯 所有參數設定完成");

            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ 設定面積參數失敗: {ex.Message}");
                Debug.WriteLine($"❌ 錯誤堆疊: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 設定厚度參數
        /// </summary>
        private void SetThicknessParameter(DirectShape element, double thicknessMm)
        {
            try
            {
                // 轉換為 Revit 內部單位 (英尺)
                double thicknessFt = thicknessMm / 304.8; // mm → ft
                
                // 嘗試設定 "厚度" 共用參數
                var thicknessParam = element.LookupParameter("厚度");
                if (thicknessParam != null && !thicknessParam.IsReadOnly)
                {
                    thicknessParam.Set(thicknessFt);
                    Debug.WriteLine($"✅ 成功設定厚度參數: {thicknessMm:F1} mm ({thicknessFt:F6} ft)");
                }
                else
                {
                    // 嘗試內建厚度參數
                    var builtInThickness = element.get_Parameter(BuiltInParameter.GENERIC_THICKNESS);
                    if (builtInThickness != null && !builtInThickness.IsReadOnly)
                    {
                        builtInThickness.Set(thicknessFt);
                        Debug.WriteLine($"✅ 成功設定內建厚度參數: {thicknessMm:F1} mm");
                    }
                    else
                    {
                        Debug.WriteLine($"⚠️ 找不到可用的厚度參數");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ 設定厚度參數失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 設定材質參數
        /// </summary>
        private void SetMaterialParameter(DirectShape element, Material material)
        {
            try
            {
                if (material == null)
                {
                    Debug.WriteLine($"⚠️ 材質為空，跳過設定");
                    return;
                }

                // 設定 MATERIAL_ID_PARAM
                var materialParam = element.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM);
                if (materialParam != null && !materialParam.IsReadOnly)
                {
                    materialParam.Set(material.Id);
                    Debug.WriteLine($"✅ 成功設定材質參數: {material.Name} (ID: {material.Id})");
                }
                else
                {
                    Debug.WriteLine($"⚠️ 無法設定材質參數");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ 設定材質參數失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 創建即時預覽元素
        /// </summary>
        private ElementId CreatePreviewElement(Document doc, Element hostElement, PlanarFace face, double thickness)
        {
            try
            {
                // 簡化的預覽版本，只做基本的面擠出
                var normal = face.FaceNormal;
                var curveLoops = face.GetEdgesAsCurveLoops();
                if (curveLoops.Count == 0) return ElementId.InvalidElementId;

                var thicknessInFeet = thickness / 304.8;
                var extrusionVector = normal.Multiply(thicknessInFeet);
                var previewSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                    curveLoops, extrusionVector, thicknessInFeet);

                if (previewSolid?.Volume <= 1e-6) return ElementId.InvalidElementId;

                // 建立預覽 DirectShape
                var directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                directShape.ApplicationId = "YD_BIM_Formwork_Preview";
                directShape.ApplicationDataId = "Preview";
                directShape.SetShape(new GeometryObject[] { previewSolid });
                directShape.Name = $"預覽模板_{hostElement.Id}";

                // 設定預覽樣式（半透明、虛線）
                var overrides = new OverrideGraphicSettings();
                overrides.SetSurfaceTransparency(50); // 50% 透明度
                overrides.SetProjectionLineWeight(1);
                overrides.SetProjectionLinePatternId(LinePatternElement.GetSolidPatternId());
                
                if (_currentMaterial?.Color.IsValid == true)
                {
                    overrides.SetProjectionLineColor(_currentMaterial.Color);
                    overrides.SetSurfaceForegroundPatternColor(_currentMaterial.Color);
                }

                var view = doc.ActiveView;
                if (view != null)
                {
                    view.SetElementOverrides(directShape.Id, overrides);
                }

                return directShape.Id;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"創建預覽元素失敗: {ex.Message}");
                return ElementId.InvalidElementId;
            }
        }

        /// <summary>
        /// 清理所有預覽元素
        /// </summary>
        private void ClearPreviewElements(Document doc)
        {
            try
            {
                if (_previewElements.Count > 0)
                {
                    using (var t = new SubTransaction(doc))
                    {
                        t.Start();
                        foreach (var id in _previewElements)
                        {
                            try
                            {
                                doc.Delete(id);
                            }
                            catch
                            {
                                // 忽略刪除失敗
                            }
                        }
                        t.Commit();
                    }
                    _previewElements.Clear();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"清理預覽元素失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 直接為選中的面生成模板（不使用傳統架構的全面掃描）
        /// </summary>
        private ElementId CreateSingleFaceFormworkDirectly(Document doc, Element hostElement, PlanarFace face, double thicknessMm, Material material)
        {
            try
            {
                Debug.WriteLine($"為面生成模板：面積 {face.Area * 0.092903:F2} m²");
                
                // 檢查面積
                var areaM2 = face.Area * 0.092903;
                if (areaM2 < 0.01)
                {
                    Debug.WriteLine("面積過小，跳過");
                    return ElementId.InvalidElementId;
                }

                // 從面建立基本模板實體
                var thickness = thicknessMm / 304.8; // 轉換為英尺
                var normal = face.FaceNormal;
                var curveLoops = face.GetEdgesAsCurveLoops();
                if (curveLoops.Count == 0)
                {
                    Debug.WriteLine("無法取得面的邊界");
                    return ElementId.InvalidElementId;
                }

                // 創建向外的擠出向量
                var extrusionVector = normal.Multiply(thickness);
                var formworkSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                    curveLoops, extrusionVector, thickness);

                if (formworkSolid?.Volume <= 1e-6)
                {
                    Debug.WriteLine("生成的模板實體體積過小");
                    return ElementId.InvalidElementId;
                }

                // 🔧 修正: 獲取附近結構元素進行接觸面扣除
                // 使用與結構分析相同的搜尋邏輯: 模板厚度 + 3000mm 緩衝 (確保能找到上方樓板)
                var searchRadiusMm = thicknessMm + 3000.0;
                Debug.WriteLine($"🔍 搜尋半徑: {searchRadiusMm:F0}mm ({searchRadiusMm/304.8:F2}ft)");
                
                var nearbyElements = GeometryExtractor.GetNearbyStructuralElementsFromFace(doc, hostElement, face, searchRadiusMm);
                
                // 🔧 改進: 根據宿主元素類型調整扣除策略
                // 柱子穿過樓板時,需要更積極的扣除策略 (降低閾值)
                bool isColumn = hostElement.Category?.Id?.Value == (long)BuiltInCategory.OST_StructuralColumns;
                double intersectionThreshold = isColumn ? 0.01 : 0.05; // 柱子使用 1% 閾值,其他使用 5%
                
                Debug.WriteLine($"🎯 宿主元素類型: {hostElement.Category?.Name}, 使用閾值: {intersectionThreshold:F2} ({intersectionThreshold * 100}%)");
                
                // 進行徹底的接觸面扣除（使用統一工具類別方法，並傳入宿主元素）
                Solid finalFormwork = GeometryExtractor.ApplySmartContactDeduction(formworkSolid, nearbyElements, intersectionThreshold, hostElement);

                if (finalFormwork?.Volume <= 1e-6)
                {
                    Debug.WriteLine("接觸面扣除後體積過小");
                    return ElementId.InvalidElementId;
                }

                // 建立 DirectShape
                var directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                directShape.ApplicationId = "YD_BIM_Formwork";
                directShape.ApplicationDataId = "SingleFace";
                directShape.SetShape(new GeometryObject[] { finalFormwork });
                directShape.Name = $"面選模板_{hostElement.Id}_{DateTime.Now:HHmmss}";

                // 設定材料（通過參數設定）
                if (material?.Id != null && material.Id != ElementId.InvalidElementId)
                {
                    try
                    {
                        var materialParam = directShape.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM);
                        if (materialParam != null && !materialParam.IsReadOnly)
                        {
                            materialParam.Set(material.Id);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"設定材料失敗: {ex.Message}");
                    }
                }

                // 設定共用參數 - 宿主元素ID
                try
                {
                    var hostIdParam = directShape.LookupParameter(SharedParams.P_HostId);
                    if (hostIdParam != null && !hostIdParam.IsReadOnly)
                    {
                        hostIdParam.Set(hostElement.Id.Value.ToString());
                        Debug.WriteLine($"✅ 設定宿主ID參數: {hostElement.Id.Value}");
                    }
                    else
                    {
                        Debug.WriteLine($"⚠️ 無法設定宿主ID參數");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"設定宿主ID參數失敗: {ex.Message}");
                }

                Debug.WriteLine($"成功生成面選模板 ID: {directShape.Id.Value}，體積: {finalFormwork.Volume:F6}");
                return directShape.Id;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"單面模板生成失敗: {ex.Message}");
                return ElementId.InvalidElementId;
            }
        }

        /// <summary>
        /// 獲取元素的所有實體（已整合到 GeometryExtractor，保留用於向後兼容）
        /// </summary>
        private List<Solid> GetElementSolids(Element element)
        {
            var solids = new List<Solid>();
            try
            {
                var geometry = element.get_Geometry(new Options { DetailLevel = ViewDetailLevel.Fine });
                if (geometry == null) return solids;

                foreach (var geomObj in geometry)
                {
                    if (geomObj is Solid solid && solid.Volume > 1e-6)
                    {
                        solids.Add(solid);
                    }
                    else if (geomObj is GeometryInstance instance)
                    {
                        var instGeom = instance.GetInstanceGeometry();
                        foreach (var instObj in instGeom)
                        {
                            if (instObj is Solid instSolid && instSolid.Volume > 1e-6)
                            {
                                solids.Add(instSolid);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"獲取元素 {element.Id} 的實體失敗: {ex.Message}");
            }
            return solids;
        }

        /// <summary>
        /// 顯示備用統計摘要（當主要對話框失敗時使用）
        /// </summary>
        private void ShowFallbackSummary(int created, double totalAreaM2)
        {
            try
            {
                Debug.WriteLine($"🔍 顯示備用統計摘要，created={created}, totalAreaM2={totalAreaM2:F6}");
                
                // 計算統計數據與主要對話框保持一致
                var totalCount = _createdFormworkItems.Count;
                var calculatedTotalArea = _createdFormworkItems.Sum(item => item.Area);
                
                // 使用實際數據，而不是可能不準確的變數
                var finalCount = Math.Max(created, totalCount);
                var finalArea = Math.Max(totalAreaM2, calculatedTotalArea);
                
                Debug.WriteLine($"📊 最終統計: 數量={finalCount}, 面積={finalArea:F6} m²");
                
                var summaryMsg = $"✅ 面選模板完成\n\n" +
                               $"📊 統計資訊:\n" +
                               $"• 模板數量: {finalCount} 個\n" +
                               $"• 總面積: {finalArea:F2} m²\n" +
                               $"• 材料: {_currentMaterial?.Name ?? "預設"}\n" +
                               $"• 厚度: {_currentThickness} mm";

                // 添加詳細項目信息
                if (_createdFormworkItems.Any())
                {
                    summaryMsg += "\n\n📋 詳細項目:";
                    foreach (var item in _createdFormworkItems.Take(5)) // 最多顯示5個項目
                    {
                        summaryMsg += $"\n• ID:{item.ElementId.Value}, 面積:{item.Area:F2} m²";
                    }
                    if (_createdFormworkItems.Count > 5)
                    {
                        summaryMsg += $"\n... 還有 {_createdFormworkItems.Count - 5} 個項目";
                    }
                }

                Debug.WriteLine($"📋 備用對話框內容準備完成");
                TaskDialog.Show("面選模板", summaryMsg);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ 顯示備用統計摘要失敗: {ex.Message}");
                // 最後的救援方案 - 最簡單的顯示
                TaskDialog.Show("面選模板", $"模板建立完成: {created} 個");
            }
        }

        /// <summary>
        /// 根據宿主元素類型判斷模板類別 (統一邏輯)
        /// </summary>
        private string GetFormworkCategoryFromHost(Element hostElement)
        {
            if (hostElement == null || hostElement.Category == null)
                return "其他";

            var categoryId = hostElement.Category.Id.Value;

            if (categoryId == (long)BuiltInCategory.OST_StructuralColumns)
                return "柱模板";
            else if (categoryId == (long)BuiltInCategory.OST_StructuralFraming)
                return "梁模板";
            else if (categoryId == (long)BuiltInCategory.OST_Floors)
                return "板模板";
            else if (categoryId == (long)BuiltInCategory.OST_Walls)
                return "牆模板";
            else if (categoryId == (long)BuiltInCategory.OST_StructuralFoundation)
                return "基礎模板";
            else if (categoryId == (long)BuiltInCategory.OST_Stairs)
                return "樓梯模板";
            else
                return "其他";
        }
    }
}
