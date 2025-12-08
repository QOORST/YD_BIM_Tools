using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Autodesk.Revit.DB;

// WPF 別名
using WpfWindow = System.Windows.Window;
using WpfGrid = System.Windows.Controls.Grid;
using WpfStackPanel = System.Windows.Controls.StackPanel;
using WpfLabel = System.Windows.Controls.Label;
using WpfButton = System.Windows.Controls.Button;
using WpfTextBlock = System.Windows.Controls.TextBlock;
using WpfScrollViewer = System.Windows.Controls.ScrollViewer;
using WpfThickness = System.Windows.Thickness;
using WpfOrientation = System.Windows.Controls.Orientation;
using WpfRowDef = System.Windows.Controls.RowDefinition;
using WpfColumnDef = System.Windows.Controls.ColumnDefinition;

// 避免命名衝突
using WpfColor = System.Windows.Media.Color;
using WpfColors = System.Windows.Media.Colors;

namespace YD_RevitTools.LicenseManager.Helpers.AR
{
    /// <summary>
    /// 面選模板數量計算頁面
    /// </summary>
    public class FormworkQuantityDialog : WpfWindow
    {
        private readonly List<FormworkItem> _formworkItems;
        private readonly Document _document;

        public FormworkQuantityDialog(Document doc, List<FormworkItem> items)
        {
            _document = doc;
            _formworkItems = items ?? new List<FormworkItem>();
            
            InitializeWindow();
            BuildContent();
        }

        private void InitializeWindow()
        {
            Title = "面選模板 - 數量計算統計";
            Width = 800;
            Height = 600;
            WindowStyle = WindowStyle.ToolWindow;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            FontFamily = new FontFamily("Microsoft JhengHei UI");
            FontSize = 12;
            Background = new SolidColorBrush(WpfColor.FromRgb(240, 240, 240));
        }

        private void BuildContent()
        {
            var mainGrid = new WpfGrid { Margin = new WpfThickness(15) };
            Content = mainGrid;

            // 定義行
            mainGrid.RowDefinitions.Add(new WpfRowDef { Height = GridLength.Auto }); // 標題
            mainGrid.RowDefinitions.Add(new WpfRowDef { Height = GridLength.Auto }); // 統計摘要
            mainGrid.RowDefinitions.Add(new WpfRowDef { Height = new GridLength(1, GridUnitType.Star) }); // 詳細列表
            mainGrid.RowDefinitions.Add(new WpfRowDef { Height = GridLength.Auto }); // 按鈕

            // 標題
            var titleLabel = new WpfLabel
            {
                Content = "🎯 面選模板數量計算統計",
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(WpfColor.FromRgb(34, 139, 34)),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new WpfThickness(0, 0, 0, 15)
            };
            mainGrid.Children.Add(titleLabel);
            WpfGrid.SetRow(titleLabel, 0);

            // 統計摘要
            BuildSummarySection(mainGrid);

            // 詳細列表
            BuildDetailSection(mainGrid);

            // 按鈕區域
            BuildButtonSection(mainGrid);
        }

        private void BuildSummarySection(WpfGrid mainGrid)
        {
            var summaryPanel = new WpfStackPanel
            {
                Orientation = WpfOrientation.Vertical,
                Background = new SolidColorBrush(WpfColors.White),
                Margin = new WpfThickness(0, 0, 0, 15)
            };

            var border = new System.Windows.Controls.Border
            {
                Child = summaryPanel,
                BorderBrush = new SolidColorBrush(WpfColor.FromRgb(200, 200, 200)),
                BorderThickness = new WpfThickness(1),
                CornerRadius = new System.Windows.CornerRadius(5),
                Padding = new WpfThickness(15)
            };

            // 計算統計數據
            var totalCount = _formworkItems.Count;
            var totalArea = _formworkItems.Sum(item => item.Area);
            var materialGroups = _formworkItems.GroupBy(item => item.MaterialName).ToList();

            // 除錯資訊
            System.Diagnostics.Debug.WriteLine($"🔍 FormworkQuantityDialog 除錯資訊:");
            System.Diagnostics.Debug.WriteLine($"📊 總項目數: {totalCount}");
            System.Diagnostics.Debug.WriteLine($"📏 總面積: {totalArea:F6} m²");
            
            for (int i = 0; i < _formworkItems.Count; i++)
            {
                var item = _formworkItems[i];
                System.Diagnostics.Debug.WriteLine($"📋 項目 {i + 1}: ID={item.ElementId}, 材質={item.MaterialName}, 面積={item.Area:F6} m²");
            }

            // 總計信息
            var totalInfo = new WpfStackPanel { Orientation = WpfOrientation.Horizontal, Margin = new WpfThickness(0, 0, 0, 10) };
            totalInfo.Children.Add(new WpfLabel { Content = "📊 ", FontSize = 16 });
            totalInfo.Children.Add(new WpfLabel { Content = $"模板總數量: {totalCount} 個", FontWeight = FontWeights.Bold, FontSize = 14 });
            totalInfo.Children.Add(new WpfLabel { Content = $" | 總面積: {totalArea:F2} m²", FontWeight = FontWeights.Bold, FontSize = 14, Foreground = new SolidColorBrush(WpfColor.FromRgb(220, 20, 60)) });
            summaryPanel.Children.Add(totalInfo);

            // 材質分組統計
            if (materialGroups.Any())
            {
                var materialHeader = new WpfLabel 
                { 
                    Content = "🎨 材質分組統計:", 
                    FontWeight = FontWeights.Bold, 
                    Margin = new WpfThickness(0, 5, 0, 5) 
                };
                summaryPanel.Children.Add(materialHeader);

                foreach (var group in materialGroups.OrderByDescending(g => g.Sum(item => item.Area)))
                {
                    var groupCount = group.Count();
                    var groupArea = group.Sum(item => item.Area);
                    var groupInfo = new WpfStackPanel { Orientation = WpfOrientation.Horizontal, Margin = new WpfThickness(20, 2, 0, 2) };
                    
                    groupInfo.Children.Add(new WpfLabel { Content = "▸", Foreground = new SolidColorBrush(WpfColor.FromRgb(100, 100, 100)) });
                    groupInfo.Children.Add(new WpfLabel { Content = $"{group.Key}: {groupCount} 個", Width = 200 });
                    groupInfo.Children.Add(new WpfLabel { Content = $"面積: {groupArea:F2} m²", Foreground = new SolidColorBrush(WpfColor.FromRgb(34, 139, 34)) });
                    
                    summaryPanel.Children.Add(groupInfo);
                }
            }

            mainGrid.Children.Add(border);
            WpfGrid.SetRow(border, 1);
        }

        private void BuildDetailSection(WpfGrid mainGrid)
        {
            var detailGroup = new System.Windows.Controls.GroupBox
            {
                Header = "📋 詳細項目列表",
                Margin = new WpfThickness(0, 0, 0, 15),
                FontWeight = FontWeights.Bold
            };

            var scrollViewer = new WpfScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            var detailPanel = new WpfStackPanel
            {
                Orientation = WpfOrientation.Vertical,
                Background = new SolidColorBrush(WpfColors.White)
            };

            // 表頭
            var headerPanel = new WpfGrid
            {
                Background = new SolidColorBrush(WpfColor.FromRgb(240, 248, 255)),
                Margin = new WpfThickness(0, 0, 0, 1)
            };
            headerPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(60) });  // 編號
            headerPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(120) }); // 元素ID
            headerPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(180) }); // 材質名稱
            headerPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(100) }); // 厚度
            headerPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(120) }); // 面積
            headerPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(1, GridUnitType.Star) }); // 備註

            var headers = new[] { "編號", "元素ID", "材質名稱", "厚度(mm)", "面積(m²)", "備註" };
            for (int i = 0; i < headers.Length; i++)
            {
                var label = new WpfLabel
                {
                    Content = headers[i],
                    FontWeight = FontWeights.Bold,
                    BorderBrush = new SolidColorBrush(WpfColor.FromRgb(200, 200, 200)),
                    BorderThickness = new WpfThickness(0, 0, 1, 1),
                    Padding = new WpfThickness(8, 5, 8, 5),
                    Background = new SolidColorBrush(WpfColor.FromRgb(230, 240, 250))
                };
                headerPanel.Children.Add(label);
                WpfGrid.SetColumn(label, i);
            }
            detailPanel.Children.Add(headerPanel);

            // 數據行
            for (int idx = 0; idx < _formworkItems.Count; idx++)
            {
                var item = _formworkItems[idx];
                var rowPanel = new WpfGrid
                {
                    Background = idx % 2 == 0 ? new SolidColorBrush(WpfColors.White) : new SolidColorBrush(WpfColor.FromRgb(248, 248, 248)),
                    Margin = new WpfThickness(0, 0, 0, 1)
                };

                // 使用相同的列定義
                rowPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(60) });
                rowPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(120) });
                rowPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(180) });
                rowPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(100) });
                rowPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(120) });
                rowPanel.ColumnDefinitions.Add(new WpfColumnDef { Width = new GridLength(1, GridUnitType.Star) });

                var values = new[]
                {
                    (idx + 1).ToString(),
                    item.ElementId.ToString(),
                    item.MaterialName,
                    item.Thickness.ToString("F1"),
                    item.Area.ToString("F2"),
                    item.Notes
                };

                for (int i = 0; i < values.Length; i++)
                {
                    var label = new WpfLabel
                    {
                        Content = values[i],
                        BorderBrush = new SolidColorBrush(WpfColor.FromRgb(220, 220, 220)),
                        BorderThickness = new WpfThickness(0, 0, 1, 1),
                        Padding = new WpfThickness(8, 5, 8, 5),
                        VerticalAlignment = VerticalAlignment.Center
                    };

                    // 面積列使用不同顏色
                    if (i == 4)
                    {
                        label.Foreground = new SolidColorBrush(WpfColor.FromRgb(34, 139, 34));
                        label.FontWeight = FontWeights.SemiBold;
                    }

                    rowPanel.Children.Add(label);
                    WpfGrid.SetColumn(label, i);
                }

                detailPanel.Children.Add(rowPanel);
            }

            scrollViewer.Content = detailPanel;
            detailGroup.Content = scrollViewer;
            mainGrid.Children.Add(detailGroup);
            WpfGrid.SetRow(detailGroup, 2);
        }

        private void BuildButtonSection(WpfGrid mainGrid)
        {
            var buttonPanel = new WpfStackPanel
            {
                Orientation = WpfOrientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new WpfThickness(0, 10, 0, 0)
            };

            var exportButton = new WpfButton
            {
                Content = "📤 匯出 CSV",
                Width = 120,
                Height = 36,
                Margin = new WpfThickness(0, 0, 10, 0),
                FontSize = 14,
                Background = new SolidColorBrush(WpfColor.FromRgb(70, 130, 180)),
                Foreground = new SolidColorBrush(WpfColors.White),
                BorderBrush = new SolidColorBrush(WpfColor.FromRgb(0, 90, 140))
            };

            var closeButton = new WpfButton
            {
                Content = "關閉",
                Width = 120,
                Height = 36,
                IsCancel = true,
                FontSize = 14,
                Background = new SolidColorBrush(WpfColor.FromRgb(220, 220, 220)),
                Foreground = new SolidColorBrush(WpfColor.FromRgb(64, 64, 64)),
                BorderBrush = new SolidColorBrush(WpfColor.FromRgb(160, 160, 160))
            };

            exportButton.Click += OnExportClick;
            closeButton.Click += (s, e) => Close();

            buttonPanel.Children.Add(exportButton);
            buttonPanel.Children.Add(closeButton);

            mainGrid.Children.Add(buttonPanel);
            WpfGrid.SetRow(buttonPanel, 3);
        }

        private void OnExportClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    DefaultExt = "csv",
                    FileName = $"面選模板統計_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    ExportToCsv(saveDialog.FileName);
                    MessageBox.Show($"成功匯出至: {saveDialog.FileName}", "匯出完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"匯出失敗: {ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportToCsv(string filePath)
        {
            using (var writer = new System.IO.StreamWriter(filePath, false, System.Text.Encoding.UTF8))
            {
                // 寫入 BOM 以確保 Excel 正確識別 UTF-8
                writer.WriteLine("編號,元素ID,材質名稱,厚度(mm),面積(m²),備註");

                for (int i = 0; i < _formworkItems.Count; i++)
                {
                    var item = _formworkItems[i];
                    writer.WriteLine($"{i + 1},{item.ElementId},{item.MaterialName},{item.Thickness:F1},{item.Area:F2},\"{item.Notes}\"");
                }

                // 寫入統計摘要
                writer.WriteLine();
                writer.WriteLine("統計摘要");
                writer.WriteLine($"總數量,{_formworkItems.Count}");
                writer.WriteLine($"總面積(m²),{_formworkItems.Sum(item => item.Area):F2}");
                
                var materialGroups = _formworkItems.GroupBy(item => item.MaterialName);
                foreach (var group in materialGroups)
                {
                    writer.WriteLine($"{group.Key}數量,{group.Count()}");
                    writer.WriteLine($"{group.Key}面積(m²),{group.Sum(item => item.Area):F2}");
                }
            }
        }
    }

    /// <summary>
    /// 模板項目數據類別
    /// </summary>
    public class FormworkItem
    {
        public ElementId ElementId { get; set; }
        public string MaterialName { get; set; }
        public double Thickness { get; set; }
        public double Area { get; set; }
        public string Notes { get; set; }

        public FormworkItem()
        {
            ElementId = ElementId.InvalidElementId;
            MaterialName = "預設";
            Thickness = 0;
            Area = 0;
            Notes = "";
        }
    }
}