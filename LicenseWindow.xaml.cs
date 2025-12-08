// LicenseWindow.xaml.cs
using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;

namespace YD_RevitTools.LicenseManager.UI
{
    public partial class LicenseWindow : Window
    {
        public LicenseWindow()
        {
            InitializeComponent();
            LoadLicenseInfo();
        }

        private void LoadLicenseInfo()
        {
            var result = LicenseManager.Instance.ValidateLicense();

            if (result.IsValid && result.LicenseInfo != null)
            {
                var license = result.LicenseInfo;

                // 設定狀態
                statusIndicator.Fill = GetStatusColor(result.Severity);
                txtStatus.Text = result.Severity == ValidationSeverity.Warning ? "已啟用（即將到期）" : "已啟用";
                txtStatus.Foreground = new SolidColorBrush(GetStatusColorValue(result.Severity));

                // 顯示授權資訊
                txtLicenseType.Text = license.GetLicenseTypeName();
                txtUserName.Text = license.UserName;
                txtCompany.Text = license.Company;
                txtStartDate.Text = license.StartDate.ToString("yyyy-MM-dd");
                txtExpiryDate.Text = license.ExpiryDate.ToString("yyyy-MM-dd");

                // 顯示剩餘天數
                txtDaysRemaining.Text = $"(剩餘 {result.DaysUntilExpiry} 天)";
                txtDaysRemaining.Foreground = result.Severity == ValidationSeverity.Warning
                    ? new SolidColorBrush(Colors.Orange)
                    : new SolidColorBrush(Color.FromRgb(102, 102, 102));

                // 設定授權類型徽章
                SetLicenseTypeBadge(license.LicenseType);

                // 顯示機器碼
                txtMachineCode.Text = MachineCodeHelper.GetMachineCode();

                // 顯示授權金鑰（遮蔽部分）
                if (!string.IsNullOrEmpty(license.LicenseKey))
                {
                    if (license.LicenseKey.Length > 20)
                    {
                        txtLicenseKeyDisplay.Text = license.LicenseKey.Substring(0, 10) +
                            "..." + license.LicenseKey.Substring(license.LicenseKey.Length - 10);
                    }
                    else
                    {
                        txtLicenseKeyDisplay.Text = license.LicenseKey;
                    }
                }

                // 檢查警告訊息
                if (result.Severity == ValidationSeverity.Warning)
                {
                    ShowLicenseMessage(result.Message, MessageType.Warning);
                }
                else
                {
                    licenseMessageBorder.Visibility = Visibility.Collapsed;
                }

                // 顯示功能清單
                DisplayFeatures(license.LicenseType);
            }
            else
            {
                // 未啟用狀態
                statusIndicator.Fill = new SolidColorBrush(Colors.Red);
                txtStatus.Text = "未啟用";
                txtStatus.Foreground = new SolidColorBrush(Colors.Red);

                txtLicenseType.Text = "-";
                txtUserName.Text = "-";
                txtCompany.Text = "-";
                txtStartDate.Text = "-";
                txtExpiryDate.Text = "-";
                txtDaysRemaining.Text = "";
                txtMachineCode.Text = MachineCodeHelper.GetMachineCode();
                txtLicenseKeyDisplay.Text = "-";

                licenseTypeBadge.Visibility = Visibility.Collapsed;

                ShowLicenseMessage(result.Message, MessageType.Error);

                featureList.Children.Clear();
            }
        }

        private void SetLicenseTypeBadge(LicenseType licenseType)
        {
            licenseTypeBadge.Visibility = Visibility.Visible;

            switch (licenseType)
            {
                case LicenseType.Trial:
                    licenseTypeBadge.Background = new SolidColorBrush(Color.FromRgb(255, 152, 0)); // Orange
                    txtLicenseTypeBadge.Text = "30天";
                    break;

                case LicenseType.Standard:
                    licenseTypeBadge.Background = new SolidColorBrush(Color.FromRgb(33, 150, 243)); // Blue
                    txtLicenseTypeBadge.Text = "365天";
                    break;

                case LicenseType.Professional:
                    licenseTypeBadge.Background = new SolidColorBrush(Color.FromRgb(76, 175, 80)); // Green
                    txtLicenseTypeBadge.Text = "365天";
                    break;
            }
        }

        private void DisplayFeatures(LicenseType licenseType)
        {
            featureList.Children.Clear();

            switch (licenseType)
            {
                case LicenseType.Trial:
                    AddFeature("✓ 基本功能", true);
                    AddFeature("✗ 標準功能", false);
                    AddFeature("✗ 進階功能", false);
                    break;

                case LicenseType.Standard:
                    AddFeature("✓ 基本功能", true);
                    AddFeature("✓ 標準功能", true);
                    AddFeature("✗ 進階功能", false);
                    break;

                case LicenseType.Professional:
                    AddFeature("✓ 基本功能", true);
                    AddFeature("✓ 標準功能", true);
                    AddFeature("✓ 進階功能", true);
                    break;
            }
        }

        private void AddFeature(string featureName, bool isAvailable)
        {
            var textBlock = new TextBlock
            {
                Text = featureName,
                Margin = new Thickness(0, 3, 0, 3),
                Foreground = new SolidColorBrush(isAvailable ? Colors.Green : Color.FromRgb(158, 158, 158))
            };
            featureList.Children.Add(textBlock);
        }

        private void ShowLicenseMessage(string message, MessageType type)
        {
            licenseMessageBorder.Visibility = Visibility.Visible;

            switch (type)
            {
                case MessageType.Success:
                    licenseMessageBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(129, 199, 132));
                    licenseMessageBorder.Background = new SolidColorBrush(Color.FromRgb(232, 245, 233));
                    txtLicenseMessage.Foreground = new SolidColorBrush(Color.FromRgb(56, 142, 60));
                    break;

                case MessageType.Warning:
                    licenseMessageBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(255, 167, 38));
                    licenseMessageBorder.Background = new SolidColorBrush(Color.FromRgb(255, 243, 224));
                    txtLicenseMessage.Foreground = new SolidColorBrush(Color.FromRgb(230, 81, 0));
                    break;

                case MessageType.Error:
                    licenseMessageBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(229, 115, 115));
                    licenseMessageBorder.Background = new SolidColorBrush(Color.FromRgb(255, 235, 238));
                    txtLicenseMessage.Foreground = new SolidColorBrush(Color.FromRgb(198, 40, 40));
                    break;
            }

            txtLicenseMessage.Text = message;
        }

        private SolidColorBrush GetStatusColor(ValidationSeverity severity)
        {
            switch (severity)
            {
                case ValidationSeverity.Success:
                    return new SolidColorBrush(Colors.Green);
                case ValidationSeverity.Warning:
                    return new SolidColorBrush(Colors.Orange);
                case ValidationSeverity.Error:
                    return new SolidColorBrush(Colors.Red);
                default:
                    return new SolidColorBrush(Colors.Gray);
            }
        }

        private Color GetStatusColorValue(ValidationSeverity severity)
        {
            switch (severity)
            {
                case ValidationSeverity.Success:
                    return Colors.Green;
                case ValidationSeverity.Warning:
                    return Colors.Orange;
                case ValidationSeverity.Error:
                    return Colors.Red;
                default:
                    return Colors.Gray;
            }
        }

        private void BtnActivate_Click(object sender, RoutedEventArgs e)
        {
            string key = txtLicenseKey.Text.Trim();

            if (string.IsNullOrEmpty(key))
            {
                MessageBox.Show("請輸入授權金鑰", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var license = ParseLicenseKey(key);

                if (LicenseManager.Instance.SaveLicense(license))
                {
                    MessageBox.Show(
                        $"授權啟用成功！\n\n" +
                        $"授權類型：{license.GetLicenseTypeName()}\n" +
                        $"使用者：{license.UserName}\n" +
                        $"到期日期：{license.ExpiryDate:yyyy-MM-dd}",
                        "成功",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    txtLicenseKey.Clear();
                    LoadLicenseInfo();
                }
                else
                {
                    MessageBox.Show("授權儲存失敗，請重試", "錯誤",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"授權金鑰無效：{ex.Message}", "錯誤",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private LicenseInfo ParseLicenseKey(string key)
        {
            try
            {
                byte[] data = Convert.FromBase64String(key);
                string json = System.Text.Encoding.UTF8.GetString(data);
                var license = Newtonsoft.Json.JsonConvert.DeserializeObject<LicenseInfo>(json);

                if (license == null)
                    throw new Exception("無法解析授權資訊");

                // 驗證授權資訊的有效性
                if (string.IsNullOrEmpty(license.UserName))
                    throw new Exception("授權資訊不完整：缺少使用者名稱");

                if (string.IsNullOrEmpty(license.Company))
                    throw new Exception("授權資訊不完整：缺少公司名稱");

                if (license.ExpiryDate <= license.StartDate)
                    throw new Exception("授權日期設定錯誤");

                return license;
            }
            catch (FormatException)
            {
                throw new Exception("授權金鑰格式錯誤");
            }
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            txtLicenseKey.Clear();
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LicenseManager.Instance.ReloadLicense();
            LoadLicenseInfo();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnCopyMachineCode_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string machineCode = txtMachineCode.Text;
                if (!string.IsNullOrEmpty(machineCode) && machineCode != "-")
                {
                    Clipboard.SetText(machineCode);
                    MessageBox.Show("機器碼已複製到剪貼簿", "成功",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"複製失敗：{ex.Message}", "錯誤",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private enum MessageType
        {
            Success,
            Warning,
            Error
        }
    }
}