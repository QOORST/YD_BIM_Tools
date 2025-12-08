using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace YD_RevitTools.LicenseManager.Helpers.AR
{
    /// <summary>
    /// 帶即時預覽的模板面選取器
    /// </summary>
    public class FormworkPreviewSelector : ISelectionFilter
    {
        private readonly Document _doc;
        private readonly UIDocument _uidoc;
        private readonly Material _material;
        private readonly double _thickness;
        private readonly HashSet<string> _selectedFaces;
        private ElementId _currentPreviewId = ElementId.InvalidElementId;

        public FormworkPreviewSelector(Document doc, UIDocument uidoc, Material material, 
            double thickness, HashSet<string> selectedFaces)
        {
            _doc = doc;
            _uidoc = uidoc;
            _material = material;
            _thickness = thickness;
            _selectedFaces = selectedFaces;
        }

        public bool AllowElement(Element elem)
        {
            // 允許結構元素（梁、柱、牆、樓板等）
            return elem.Category?.Id.Value == (long)BuiltInCategory.OST_StructuralFraming ||
                   elem.Category?.Id.Value == (long)BuiltInCategory.OST_StructuralColumns ||
                   elem.Category?.Id.Value == (long)BuiltInCategory.OST_Walls ||
                   elem.Category?.Id.Value == (long)BuiltInCategory.OST_Floors;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            try
            {
                var element = _doc.GetElement(reference.ElementId);
                var face = element?.GetGeometryObjectFromReference(reference) as PlanarFace;
                
                if (face == null) return false;

                // 檢查是否已選取過
                var faceKey = GetFaceKey(element, face);
                if (_selectedFaces.Contains(faceKey))
                {
                    return false; // 已選取的面不允許再次選取
                }

                // 清理之前的預覽
                ClearCurrentPreview();

                // 創建新的預覽
                CreatePreview(element, face);

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"預覽選取失敗: {ex.Message}");
                return false;
            }
        }

        private void CreatePreview(Element hostElement, PlanarFace face)
        {
            try
            {
                using (var subTransaction = new SubTransaction(_doc))
                {
                    subTransaction.Start();

                    // 創建預覽模板
                    var normal = face.FaceNormal;
                    var curveLoops = face.GetEdgesAsCurveLoops();
                    if (curveLoops.Count == 0) return;

                    // 稍微偏移以避免與原結構重疊
                    var offsetDistance = 0.01; // 1cm
                    var offsetVector = normal.Multiply(offsetDistance);
                    
                    var offsetLoops = new List<CurveLoop>();
                    foreach (var loop in curveLoops)
                    {
                        var offsetLoop = new CurveLoop();
                        foreach (Curve curve in loop)
                        {
                            var offsetCurve = curve.CreateTransformed(Transform.CreateTranslation(offsetVector));
                            offsetLoop.Append(offsetCurve);
                        }
                        offsetLoops.Add(offsetLoop);
                    }

                    var thicknessInFeet = _thickness / 304.8;
                    var extrusionVector = normal.Multiply(thicknessInFeet);
                    var previewSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                        offsetLoops, extrusionVector, thicknessInFeet);

                    if (previewSolid?.Volume <= 1e-6) return;

                    // 建立預覽 DirectShape
                    var directShape = DirectShape.CreateElement(_doc, new ElementId(BuiltInCategory.OST_GenericModel));
                    directShape.ApplicationId = "YD_BIM_Formwork_Preview";
                    directShape.ApplicationDataId = "LivePreview";
                    directShape.SetShape(new GeometryObject[] { previewSolid });
                    directShape.Name = $"即時預覽_{hostElement.Id.Value}";

                    _currentPreviewId = directShape.Id;

                    // 設定預覽樣式
                    var overrides = new OverrideGraphicSettings();
                    overrides.SetSurfaceTransparency(60); // 60% 透明度
                    overrides.SetProjectionLineWeight(2);
                    
                    if (_material?.Color.IsValid == true)
                    {
                        // 使用材料顏色但稍微調亮
                        var previewColor = new Color(
                            (byte)Math.Min(255, _material.Color.Red + 50),
                            (byte)Math.Min(255, _material.Color.Green + 50),
                            (byte)Math.Min(255, _material.Color.Blue + 50)
                        );
                        overrides.SetProjectionLineColor(previewColor);
                        overrides.SetSurfaceForegroundPatternColor(previewColor);
                    }
                    else
                    {
                        // 預設預覽顏色（淺藍色）
                        overrides.SetProjectionLineColor(new Color(100, 150, 255));
                        overrides.SetSurfaceForegroundPatternColor(new Color(100, 150, 255));
                    }

                    var view = _doc.ActiveView;
                    if (view != null)
                    {
                        view.SetElementOverrides(directShape.Id, overrides);
                    }

                    subTransaction.Commit();
                    _uidoc.RefreshActiveView();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"創建預覽失敗: {ex.Message}");
            }
        }

        private void ClearCurrentPreview()
        {
            try
            {
                if (_currentPreviewId != ElementId.InvalidElementId)
                {
                    using (var subTransaction = new SubTransaction(_doc))
                    {
                        subTransaction.Start();
                        _doc.Delete(_currentPreviewId);
                        subTransaction.Commit();
                    }
                    _currentPreviewId = ElementId.InvalidElementId;
                    _uidoc.RefreshActiveView();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"清理預覽失敗: {ex.Message}");
            }
        }

        public void ClearAllPreviews()
        {
            ClearCurrentPreview();
        }

        private string GetFaceKey(Element element, PlanarFace face)
        {
            try
            {
                var origin = face.Origin;
                var normal = face.FaceNormal;
                var area = face.Area;
                
                return $"{element.Id.Value}_{origin.X:F3}_{origin.Y:F3}_{origin.Z:F3}_{normal.X:F3}_{normal.Y:F3}_{normal.Z:F3}_{area:F6}";
            }
            catch
            {
                return $"{element.Id.Value}_{DateTime.Now.Ticks}";
            }
        }
    }
}