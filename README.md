# YD_RevitTools.LicenseManager

## 專案說明
YD BIM 工具的授權管理系統，用於管理 Revit 外掛程式的授權驗證。

## 功能特色
- ✅ 支援三種授權類型：試用版(30天)、標準版(365天)、專業版(365天)
- ✅ 機器碼綁定功能
- ✅ 授權到期提醒
- ✅ 加密儲存授權資訊
- ✅ 圖形化授權管理介面

## 已修正的問題

### 1. 命名空間錯誤
- ❌ 原來: `YD_RevitTools.Tool1`
- ✅ 修正為: `YD_RevitTools.LicenseManager`

### 2. 類別名稱衝突
- ❌ 原來: `Application` (與 System 命名空間衝突)
- ✅ 修正為: `App`

### 3. 移除不存在的命名空間引用
- ❌ 移除: `using YD_RevitTools.LicenseManager.UI;`
- ✅ LicenseWindow 直接在 `YD_RevitTools.LicenseManager` 命名空間下

### 4. 新增缺少的 NuGet 套件
- ✅ 新增 Newtonsoft.Json 13.0.3
- ✅ 新增 packages.config

### 5. 新增必要的組件參考
- ✅ System.Management (用於取得硬體資訊)
- ✅ RevitAPI.dll (Revit 2024)
- ✅ RevitAPIUI.dll (Revit 2024)

## 建置專案前的準備

### 1. 安裝 NuGet 套件
在 Visual Studio 中，請使用 NuGet Package Manager 安裝：
```
Install-Package Newtonsoft.Json -Version 13.0.3
```

或者在專案上按右鍵 → 管理 NuGet 套件 → 搜尋並安裝 Newtonsoft.Json

### 2. 調整 Revit API 參考路徑
請根據你的 Revit 安裝版本，調整 .csproj 檔案中的路徑：

**Revit 2024:**
```xml
C:\Program Files\Autodesk\Revit 2024\RevitAPI.dll
C:\Program Files\Autodesk\Revit 2024\RevitAPIUI.dll
```

**Revit 2023:**
```xml
C:\Program Files\Autodesk\Revit 2023\RevitAPI.dll
C:\Program Files\Autodesk\Revit 2023\RevitAPIUI.dll
```

**Revit 2022:**
```xml
C:\Program Files\Autodesk\Revit 2022\RevitAPI.dll
C:\Program Files\Autodesk\Revit 2022\RevitAPIUI.dll
```

### 3. 多版本支援建議
如果需要支援多個 Revit 版本，建議：
1. 為每個版本建立不同的建置設定（Debug2024, Release2024 等）
2. 使用條件編譯來處理不同版本的 API 差異

## 專案結構

```
YD_RevitTools.LicenseManager/
├── App.cs                          # Revit 外掛入口點
├── LicenseManager.cs               # 授權管理核心邏輯
├── LicenseKeyGenerator.cs          # 授權金鑰生成工具
├── LicenseWindow.xaml              # 授權管理視窗介面
├── LicenseWindow.xaml.cs           # 授權管理視窗邏輯
├── MachineCodeHelper.cs            # 機器碼產生工具
├── Commands/
│   ├── LicenseManagementCommand.cs # 授權管理命令
│   └── AdvancedToolCommand.cs      # 進階功能範例(需專業版)
└── packages.config                 # NuGet 套件設定
```

## 使用方式

### 1. 載入到 Revit
編譯後，將 DLL 和相關檔案複製到 Revit 外掛目錄：
```
C:\ProgramData\Autodesk\Revit\Addins\2024\
```

### 2. 建立 .addin 檔案
建立 `YD_RevitTools.LicenseManager.addin`：
```xml
<?xml version="1.0" encoding="utf-8"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>YD BIM 工具</Name>
    <Assembly>YD_RevitTools.LicenseManager.dll</Assembly>
    <FullClassName>YD_RevitTools.LicenseManager.App</FullClassName>
    <ClientId>B3F5D2D4-9392-4A9E-9C0D-A6F5DD93FAC7</ClientId>
    <VendorId>YD</VendorId>
    <VendorDescription>YD BIM Tools</VendorDescription>
  </AddIn>
</RevitAddIns>
```

### 3. 生成授權金鑰
使用 LicenseKeyGenerator 工具生成授權金鑰：
1. 編譯並執行 LicenseKeyGenerator
2. 選擇授權類型
3. 輸入使用者資訊
4. 取得授權金鑰

### 4. 啟用授權
1. 在 Revit 中開啟 "YD_BIM 工具" 頁籤
2. 點擊 "授權管理" 按鈕
3. 貼上授權金鑰
4. 點擊 "啟用授權"

## 授權類型比較

| 功能 | 試用版 | 標準版 | 專業版 |
|------|--------|--------|--------|
| 有效期限 | 30 天 | 365 天 | 365 天 |
| 基本功能 | ✓ | ✓ | ✓ |
| 標準功能 | ✗ | ✓ | ✓ |
| 進階功能 | ✗ | ✗ | ✓ |

## 授權檔案位置
授權資訊加密儲存於：
```
%AppData%\YD\RevitTools\license.dat
```

## 注意事項
1. Revit API DLL 的 `Private` 屬性應設為 `False`，避免複製到輸出目錄
2. Newtonsoft.Json 的 `Private` 屬性應設為 `True`，確保會複製到輸出目錄
3. 授權金鑰使用 Base64 編碼的 JSON 格式
4. 機器碼基於 CPU ID 和主機板序號生成

## 開發環境
- Visual Studio 2019 或更新版本
- .NET Framework 4.8
- Revit 2022/2023/2024

## 授權
© 2025 YD BIM Tools. All rights reserved.
