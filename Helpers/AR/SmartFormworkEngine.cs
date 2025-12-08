using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Autodesk.Revit.DB;

namespace YD_RevitTools.LicenseManager.Helpers.AR
{
    /// <summary>
    /// 智能模板引擎 - 基於 PourZone 分群和空間索引的高效模板生成
    /// 實現用戶提供的理想架構：收集 → 分群 → 幾何鄰接判斷 → 生成模板
    /// </summary>
    public static class SmartFormworkEngine
    {
        private const double GEOMETRY_TOLERANCE = 1e-6;
        private const double DEFAULT_THICKNESS_MM = 18.0;
        
        /// <summary>
        /// 主要入口：分析整個項目並生成智能模板
        /// </summary>
        public static SmartFormworkResult AnalyzeAndGenerate(Document doc, double thicknessMm = DEFAULT_THICKNESS_MM)
        {
            var result = new SmartFormworkResult();
            
            try
            {
                Debug.WriteLine("=== 開始智能模板分析 ===");
                
                // 1. 收集所有結構元素（多類別）
                var elements = CollectStructuralElements(doc);
                Debug.WriteLine($"收集到 {elements.Count} 個結構元素");
                
                // 2. 依 PourZone 分群
                var groups = GroupByPourZone(elements);
                Debug.WriteLine($"分成 {groups.Count} 個澆置區域群組");
                
                // 3. 對每個群組進行分析和生成
                foreach (var group in groups)
                {
                    var groupResult = ProcessPourZoneGroup(doc, group.Key, group.Value, thicknessMm);
                    result.GroupResults[group.Key] = groupResult;
                    
                    // 累加統計
                    result.TotalElements += groupResult.ElementCount;
                    result.TotalFormworkArea += groupResult.TotalArea;
                    result.GeneratedFormworkIds.AddRange(groupResult.FormworkIds);
                }
                
                Debug.WriteLine($"=== 智能模板分析完成 ===");
                Debug.WriteLine($"總元素: {result.TotalElements}, 總面積: {result.TotalFormworkArea:F2} m²");
                
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"智能模板分析失敗: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Step 1: 收集結構元素（依類別預設邏輯，可在專案參數中覆寫）
        /// </summary>
        private static List<Element> CollectStructuralElements(Document doc)
        {
            var elements = new List<Element>();
            
            // 結構框架 (梁)
            var framingCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(IsConcreteCastInPlace);
            elements.AddRange(framingCollector);
            
            // 結構板
            var slabCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Floors)
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(IsConcreteCastInPlace);
            elements.AddRange(slabCollector);
            
            // 結構柱
            var columnCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(IsConcreteCastInPlace);
            elements.AddRange(columnCollector);
            
            // 牆/剪力牆
            var wallCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(IsConcreteCastInPlace);
            elements.AddRange(wallCollector);
            
            // 基礎/基腳/承台
            var foundationCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(IsConcreteCastInPlace);
            elements.AddRange(foundationCollector);
            
            Debug.WriteLine($"收集元素統計: 梁={framingCollector.Count()}, 板={slabCollector.Count()}, " +
                          $"柱={columnCollector.Count()}, 牆={wallCollector.Count()}, 基礎={foundationCollector.Count()}");
            
            return elements;
        }
        
        /// <summary>
        /// 判斷是否為現澆混凝土構件
        /// </summary>
        private static bool IsConcreteCastInPlace(Element element)
        {
            try
            {
                // 檢查材料是否為混凝土
                var material = GetElementMaterial(element);
                if (material?.Name?.ToLower().Contains("concrete") == true ||
                    material?.Name?.ToLower().Contains("混凝土") == true)
                {
                    return true;
                }
                
                // 檢查構件名稱
                var typeName = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString();
                if (typeName?.ToLower().Contains("concrete") == true ||
                    typeName?.Contains("混凝土") == true ||
                    typeName?.Contains("RC") == true)
                {
                    return true;
                }
                
                return false;
            }
            catch
            {
                return true; // 預設認為需要模板
            }
        }
        
        /// <summary>
        /// 取得元素主要材料
        /// </summary>
        private static Material GetElementMaterial(Element element)
        {
            try
            {
                var doc = element.Document;
                var materialIds = element.GetMaterialIds(false);
                if (materialIds.Count > 0)
                {
                    return doc.GetElement(materialIds.First()) as Material;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
        
        /// <summary>
        /// Step 2: 依 PourZone 分群
        /// </summary>
        private static Dictionary<string, List<Element>> GroupByPourZone(List<Element> elements)
        {
            var groups = new Dictionary<string, List<Element>>();
            
            foreach (var element in elements)
            {
                // 實現用戶的分群邏輯: AR_PourZone + Phase
                var pourZone = GetStr(element, "AR_PourZone") ?? "Default";
                var phaseParam = element.get_Parameter(BuiltInParameter.PHASE_CREATED);
                var phaseId = phaseParam?.AsElementId()?.Value ?? 0;
                var groupKey = $"{pourZone}|{phaseId}";
                
                if (!groups.ContainsKey(groupKey))
                {
                    groups[groupKey] = new List<Element>();
                }
                groups[groupKey].Add(element);
            }
            
            return groups;
        }
        
        /// <summary>
        /// 取得元素的字串參數值
        /// </summary>
        private static string GetStr(Element element, string paramName)
        {
            try
            {
                var param = element.LookupParameter(paramName);
                if (param != null && param.HasValue)
                {
                    return param.AsString();
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
        
        /// <summary>
        /// Step 3: 處理單一澆置區域群組
        /// </summary>
        private static PourZoneGroupResult ProcessPourZoneGroup(Document doc, string groupKey, 
            List<Element> groupElements, double thicknessMm)
        {
            var result = new PourZoneGroupResult
            {
                GroupKey = groupKey,
                ElementCount = groupElements.Count
            };
            
            Debug.WriteLine($"處理群組 {groupKey}: {groupElements.Count} 個元素");
            
            // 建立空間索引
            var spatialIndex = BuildSpatialIndex(groupElements);
            
            // 對每個元素進行幾何分析
            foreach (var element in groupElements)
            {
                try
                {
                    var elementResult = ProcessElement(doc, element, spatialIndex, thicknessMm);
                    result.ElementResults[element.Id] = elementResult;
                    result.TotalArea += elementResult.Area;
                    result.FormworkIds.AddRange(elementResult.FormworkIds);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"處理元素 {element.Id} 失敗: {ex.Message}");
                }
            }
            
            Debug.WriteLine($"群組 {groupKey} 完成: 面積 {result.TotalArea:F2} m²");
            return result;
        }
        
        /// <summary>
        /// Step 4: 建立空間索引（以 BBox 建索引）
        /// </summary>
        private static SpatialIndex BuildSpatialIndex(List<Element> elements)
        {
            var spatialIndex = new SpatialIndex();
            
            foreach (var element in elements)
            {
                try
                {
                    var solid = GetPrimarySolid(element);
                    if (solid != null)
                    {
                        var bbox = solid.GetBoundingBox();
                        spatialIndex.AddElement(element, bbox);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"建立元素 {element.Id} 空間索引失敗: {ex.Message}");
                }
            }
            
            Debug.WriteLine($"空間索引建立完成: {spatialIndex.ElementCount} 個元素");
            return spatialIndex;
        }
        
        /// <summary>
        /// 取得元素的主要實體
        /// </summary>
        private static Solid GetPrimarySolid(Element element)
        {
            try
            {
                var options = new Options
                {
                    ComputeReferences = true,
                    DetailLevel = ViewDetailLevel.Fine
                };
                
                var geometry = element.get_Geometry(options);
                if (geometry == null) return null;
                
                Solid largestSolid = null;
                double maxVolume = 0;
                
                foreach (var geomObj in geometry)
                {
                    if (geomObj is Solid solid && solid.Volume > GEOMETRY_TOLERANCE)
                    {
                        if (solid.Volume > maxVolume)
                        {
                            maxVolume = solid.Volume;
                            largestSolid = solid;
                        }
                    }
                    else if (geomObj is GeometryInstance instance)
                    {
                        var instGeom = instance.GetInstanceGeometry();
                        foreach (var instObj in instGeom)
                        {
                            if (instObj is Solid instSolid && instSolid.Volume > GEOMETRY_TOLERANCE)
                            {
                                if (instSolid.Volume > maxVolume)
                                {
                                    maxVolume = instSolid.Volume;
                                    largestSolid = instSolid;
                                }
                            }
                        }
                    }
                }
                
                return largestSolid;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"取得主要實體失敗: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Step 5: 處理單一元素的所有面
        /// </summary>
        private static ElementFormworkResult ProcessElement(Document doc, Element element, 
            SpatialIndex spatialIndex, double thicknessMm)
        {
            var result = new ElementFormworkResult();
            
            try
            {
                var solid = GetPrimarySolid(element);
                if (solid == null)
                {
                    Debug.WriteLine($"元素 {element.Id} 無有效實體");
                    return result;
                }
                
                // 遍歷所有面
                foreach (Face face in solid.Faces)
                {
                    if (!(face is PlanarFace planarFace))
                        continue;
                    
                    result.ProcessedFaces++;
                    
                    // Step 6: 類別+面向初篩
                    if (!PassCategoryFaceRule(element, planarFace))
                    {
                        result.ExcludedFaces++;
                        continue;
                    }
                    
                    // Step 7: 與同群鄰居接觸檢查
                    if (TouchNeighborFace(doc, element, planarFace, spatialIndex))
                    {
                        result.ExcludedFaces++;
                        continue;
                    }
                    
                    // Step 8: 旗標排除檢查
                    if (ExcludedByFlags(element, planarFace))
                    {
                        result.ExcludedFaces++;
                        continue;
                    }
                    
                    // Step 9: 生成模板
                    var formworkId = CreateFormworkDirectShape(doc, element, planarFace, thicknessMm);
                    if (formworkId != ElementId.InvalidElementId)
                    {
                        result.FormworkIds.Add(formworkId);
                        
                        // 計算面積 (注意轉單位 ft² → m²)
                        double areaM2 = planarFace.Area * 0.092903;
                        result.Area += areaM2;
                    }
                }
                
                Debug.WriteLine($"元素 {element.Id}: 處理 {result.ProcessedFaces} 面, " +
                              $"排除 {result.ExcludedFaces} 面, 生成 {result.FormworkIds.Count} 個模板");
                
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"處理元素 {element.Id} 失敗: {ex.Message}");
                return result;
            }
        }
        
        /// <summary>
        /// Step 6: 類別+面向規則判斷 (PassCategoryFaceRule)
        /// 根據用戶提供的規則表實現
        /// </summary>
        private static bool PassCategoryFaceRule(Element element, PlanarFace face)
        {
            try
            {
                var category = element.Category?.Id?.Value;
                var normal = face.FaceNormal;
                
                // 垂直面 (Z 分量小)
                bool isVertical = Math.Abs(normal.Z) < 0.3;
                // 水平面 (Z 分量大)
                bool isHorizontal = Math.Abs(normal.Z) > 0.7;
                // 向上面 (Z > 0)
                bool isFacingUp = normal.Z > 0.7;
                // 向下面 (Z < 0)
                bool isFacingDown = normal.Z < -0.7;
                
                switch (category)
                {
                    case (long)BuiltInCategory.OST_StructuralFraming: // 梁
                        // 兩側 + 底；頂面與樓板同澆時不算
                        return isVertical || isFacingDown;
                        
                    case (long)BuiltInCategory.OST_Floors: // 板
                        // 底面 + 周邊側緣（外露邊）
                        // 頂面（作為上模板平台不算），落地板（on-grade）之底面可不算
                        return isFacingDown || isVertical;
                        
                    case (long)BuiltInCategory.OST_StructuralColumns: // 柱
                        // 四周側面
                        // 與上/下接觸的梁、板、基礎之接觸面
                        return isVertical;
                        
                    case (long)BuiltInCategory.OST_Walls: // 牆/剪力牆
                        // 兩側面 + 外露頂頭
                        // 與板、梁、柱密貼面；擋土牆背填可依規則不算
                        return isVertical || isFacingUp;
                        
                    case (long)BuiltInCategory.OST_StructuralFoundation: // 基礎/基腳/承台
                        // 外周側面
                        // 與土、墊層接觸的底面、頂面（多為不需模板）
                        return isVertical;
                        
                    default:
                        // 未知類別，保守處理
                        return true;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"類別面向規則判斷失敗: {ex.Message}");
                return true; // 預設通過
            }
        }
        
        /// <summary>
        /// Step 7: 檢查面是否與同群鄰居接觸 (TouchNeighborFace)
        /// </summary>
        private static bool TouchNeighborFace(Document doc, Element element, PlanarFace face, SpatialIndex spatialIndex)
        {
            try
            {
                // 取得潛在鄰居
                var neighborIds = spatialIndex.GetNeighbors(element.Id, 0.1); // 10cm 容差
                
                foreach (var neighborId in neighborIds)
                {
                    var neighbor = doc.GetElement(neighborId);
                    if (neighbor == null) continue;
                    
                    var neighborSolid = GetPrimarySolid(neighbor);
                    if (neighborSolid == null) continue;
                    
                    // 檢查面是否與鄰居實體相交
                    if (FaceIntersectsSolid(face, neighborSolid))
                    {
                        Debug.WriteLine($"面與鄰居 {neighborId} 接觸，跳過模板生成");
                        return true;
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"鄰接檢查失敗: {ex.Message}");
                return false; // 預設不接觸
            }
        }
        
        /// <summary>
        /// 檢查面是否與實體相交
        /// </summary>
        private static bool FaceIntersectsSolid(PlanarFace face, Solid solid)
        {
            try
            {
                // 使用 BoundingBox 快速篩選
                var solidBounds = solid.GetBoundingBox();
                
                // 取得面的中心點和邊界進行簡化檢查
                var faceCenter = face.Evaluate(new UV(0.5, 0.5));
                
                // 檢查面中心點是否在實體的包圍盒內（簡化版本）
                if (faceCenter.X >= solidBounds.Min.X - 0.01 && faceCenter.X <= solidBounds.Max.X + 0.01 &&
                    faceCenter.Y >= solidBounds.Min.Y - 0.01 && faceCenter.Y <= solidBounds.Max.Y + 0.01 &&
                    faceCenter.Z >= solidBounds.Min.Z - 0.01 && faceCenter.Z <= solidBounds.Max.Z + 0.01)
                {
                    // 進一步檢查：使用 Boolean 操作檢測重疊
                    try
                    {
                        // 創建面的小體積實體進行測試
                        var faceLoops = face.GetEdgesAsCurveLoops();
                        if (faceLoops.Count > 0)
                        {
                            var normal = face.FaceNormal;
                            var testThickness = 0.001; // 1mm 測試厚度
                            var extrusionVector = normal.Multiply(testThickness);
                            var testSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                                faceLoops, extrusionVector, testThickness);
                            
                            if (testSolid != null)
                            {
                                var intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                                    testSolid, solid, BooleanOperationsType.Intersect);
                                return intersection?.Volume > GEOMETRY_TOLERANCE;
                            }
                        }
                    }
                    catch
                    {
                        // Boolean 操作失敗，回退到包圍盒檢查結果
                        return true;
                    }
                }
                
                return false;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// BoundingBox 相交檢查
        /// </summary>
        private static bool BoundingBoxIntersects(BoundingBoxXYZ box1, BoundingBoxXYZ box2)
        {
            return box1.Min.X <= box2.Max.X && box1.Max.X >= box2.Min.X &&
                   box1.Min.Y <= box2.Max.Y && box1.Max.Y >= box2.Min.Y &&
                   box1.Min.Z <= box2.Max.Z && box1.Max.Z >= box2.Min.Z;
        }
        
        /// <summary>
        /// Step 8: 旗標排除檢查 (ExcludedByFlags)
        /// </summary>
        private static bool ExcludedByFlags(Element element, PlanarFace face)
        {
            try
            {
                // 檢查 OnGrade 旗標（落地板底模）
                var onGrade = GetBool(element, "OnGrade");
                if (onGrade && face.FaceNormal.Z < -0.7) // 向下面
                {
                    Debug.WriteLine("OnGrade 旗標：跳過落地板底模");
                    return true;
                }
                
                // 檢查 AgainstSoil 旗標（擋土牆背填）
                var againstSoil = GetBool(element, "AgainstSoil");
                if (againstSoil)
                {
                    Debug.WriteLine("AgainstSoil 旗標：跳過擋土牆背填面");
                    return true;
                }
                
                // 檢查 Override 旗標（使用者自訂排除）
                var overrideExclude = GetBool(element, "FormworkOverride");
                if (overrideExclude)
                {
                    Debug.WriteLine("FormworkOverride 旗標：使用者自訂排除");
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"旗標排除檢查失敗: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// 取得元素的布林參數值
        /// </summary>
        private static bool GetBool(Element element, string paramName)
        {
            try
            {
                var param = element.LookupParameter(paramName);
                if (param != null && param.HasValue)
                {
                    return param.AsInteger() != 0;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Step 9: 建立模板 DirectShape (使用成功的傳統架構)
        /// </summary>
        private static ElementId CreateFormworkDirectShape(Document doc, Element hostElement, 
            PlanarFace face, double thicknessMm)
        {
            try
            {
                // 優先使用成功的傳統架構
                var info = FormworkEngine.AnalyzeHost(doc, hostElement, true);
                var ids = FormworkEngine.BuildFormworkSolids(
                    doc, hostElement, info, null, null, 
                    true, thicknessMm, 30, true);
                
                if (ids.Count > 0)
                {
                    Debug.WriteLine($"智能引擎使用傳統架構成功生成模板");
                    
                    // 設定智能引擎的標識
                    var element = doc.GetElement(ids.First());
                    if (element is DirectShape ds)
                    {
                        ds.Name = $"智能模板_{hostElement.Category?.Name}_{hostElement.Id.Value}";
                        SetFormworkParameters(ds, hostElement, face.Area * 0.092903);
                    }
                    
                    return ids.First();
                }
                
                // 如果傳統架構失敗，回退到原始方法
                Debug.WriteLine("傳統架構失敗，使用原始面擠出方法");
                return CreateFormworkDirectShapeFromFace(doc, hostElement, face, thicknessMm);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"建立模板失敗: {ex.Message}");
                return CreateFormworkDirectShapeFromFace(doc, hostElement, face, thicknessMm);
            }
        }
        
        /// <summary>
        /// 原始的面擠出方法（作為回退選項）
        /// </summary>
        private static ElementId CreateFormworkDirectShapeFromFace(Document doc, Element hostElement, 
            PlanarFace face, double thicknessMm)
        {
            try
            {
                var thickness = thicknessMm / 304.8; // 轉換為英尺
                var normal = face.FaceNormal;
                
                // 取得面的邊界曲線
                var curveLoops = face.GetEdgesAsCurveLoops();
                if (curveLoops.Count == 0) return ElementId.InvalidElementId;
                
                // 偏移避免與原結構重疊
                var offsetDistance = 0.01; // 1cm 偏移
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
                
                // 擠出模板
                var extrusionVector = normal.Multiply(thickness);
                var formworkSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                    offsetLoops, extrusionVector, thickness);
                
                if (formworkSolid?.Volume <= GEOMETRY_TOLERANCE)
                    return ElementId.InvalidElementId;
                
                // 建立 DirectShape
                var directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                directShape.ApplicationId = "YD_BIM_SmartFormwork";
                directShape.ApplicationDataId = "SmartEngine";
                directShape.SetShape(new GeometryObject[] { formworkSolid });
                directShape.Name = $"智能模板_{hostElement.Category?.Name}_{hostElement.Id.Value}";
                
                // 設定參數
                SetFormworkParameters(directShape, hostElement, face.Area * 0.092903);
                
                return directShape.Id;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"面擠出方法失敗: {ex.Message}");
                return ElementId.InvalidElementId;
            }
        }
        
        /// <summary>
        /// 設定模板參數
        /// </summary>
        private static void SetFormworkParameters(DirectShape formwork, Element hostElement, double areaM2)
        {
            try
            {
                // 設定註解
                var commentParam = formwork.LookupParameter("註解");
                if (commentParam != null && !commentParam.IsReadOnly)
                {
                    commentParam.Set($"智能模板 - 宿主:{hostElement.Id} - 面積:{areaM2:F2}m²");
                }
                
                // 設定標記
                var markParam = formwork.LookupParameter("標記");
                if (markParam != null && !markParam.IsReadOnly)
                {
                    markParam.Set($"SMART_{hostElement.Id.Value}");
                }
                
                Debug.WriteLine($"設定參數完成 - 面積:{areaM2:F2}m²");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"設定模板參數失敗: {ex.Message}");
            }
        }
    }
    
    /// <summary>
    /// 空間索引類別
    /// </summary>
    public class SpatialIndex
    {
        private readonly Dictionary<ElementId, BoundingBoxXYZ> _elementBounds = new Dictionary<ElementId, BoundingBoxXYZ>();
        
        public int ElementCount => _elementBounds.Count;
        
        public void AddElement(Element element, BoundingBoxXYZ bounds)
        {
            _elementBounds[element.Id] = bounds;
        }
        
        public List<ElementId> GetNeighbors(ElementId elementId, double tolerance = 0.1)
        {
            var neighbors = new List<ElementId>();
            
            if (!_elementBounds.TryGetValue(elementId, out var targetBounds))
                return neighbors;
            
            // 擴展包圍盒以找尋鄰居
            var expandedBounds = new BoundingBoxXYZ
            {
                Min = new XYZ(targetBounds.Min.X - tolerance, targetBounds.Min.Y - tolerance, targetBounds.Min.Z - tolerance),
                Max = new XYZ(targetBounds.Max.X + tolerance, targetBounds.Max.Y + tolerance, targetBounds.Max.Z + tolerance)
            };
            
            foreach (var kvp in _elementBounds)
            {
                if (kvp.Key == elementId) continue;
                
                if (BoundingBoxIntersects(expandedBounds, kvp.Value))
                {
                    neighbors.Add(kvp.Key);
                }
            }
            
            return neighbors;
        }
        
        private bool BoundingBoxIntersects(BoundingBoxXYZ box1, BoundingBoxXYZ box2)
        {
            return box1.Min.X <= box2.Max.X && box1.Max.X >= box2.Min.X &&
                   box1.Min.Y <= box2.Max.Y && box1.Max.Y >= box2.Min.Y &&
                   box1.Min.Z <= box2.Max.Z && box1.Max.Z >= box2.Min.Z;
        }
    }
    
    /// <summary>
    /// 智能模板分析結果
    /// </summary>
    public class SmartFormworkResult
    {
        public Dictionary<string, PourZoneGroupResult> GroupResults { get; } = new Dictionary<string, PourZoneGroupResult>();
        public List<ElementId> GeneratedFormworkIds { get; } = new List<ElementId>();
        public int TotalElements { get; set; }
        public double TotalFormworkArea { get; set; }
    }
    
    /// <summary>
    /// 澆置區域群組結果
    /// </summary>
    public class PourZoneGroupResult
    {
        public string GroupKey { get; set; }
        public int ElementCount { get; set; }
        public double TotalArea { get; set; }
        public List<ElementId> FormworkIds { get; } = new List<ElementId>();
        public Dictionary<ElementId, ElementFormworkResult> ElementResults { get; } = new Dictionary<ElementId, ElementFormworkResult>();
    }
    
    /// <summary>
    /// 單一元素模板結果
    /// </summary>
    public class ElementFormworkResult
    {
        public double Area { get; set; }
        public List<ElementId> FormworkIds { get; } = new List<ElementId>();
        public int ProcessedFaces { get; set; }
        public int ExcludedFaces { get; set; }
    }
}