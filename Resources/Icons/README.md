# COBie 外掛圖示說明

## 需要的圖示檔案

請在此資料夾放置以下 PNG 圖示檔案：

### 1. COBie 欄位管理
- `cobie_config_16.png` - 16x16 像素
- `cobie_config_32.png` - 32x32 像素
- 建議圖示：齒輪⚙️、設定圖示、表格配置圖示

### 2. COBie 匯出
- `cobie_export_16.png` - 16x16 像素
- `cobie_export_32.png` - 32x32 像素
- 建議圖示：向右箭頭➡️、匯出圖示、文件輸出圖示

### 3. COBie 匯入
- `cobie_import_16.png` - 16x16 像素
- `cobie_import_32.png` - 32x32 像素
- 建議圖示：向左箭頭⬅️、匯入圖示、文件輸入圖示

## 圖示設計建議

- **格式**：PNG 透明背景
- **尺寸**：
  - 小圖示：16x16 像素（用於選單）
  - 大圖示：32x32 像素（用於功能區按鈕）
- **風格**：
  - 簡潔清晰
  - 顏色建議：藍色系或綠色系（符合 BIM 工具風格）
  - 確保在淺色和深色背景下都清晰可見

## 快速生成圖示的方式

### 方式 1：使用線上工具
- [Flaticon](https://www.flaticon.com/) - 搜尋 "settings", "export", "import"
- [Icon8](https://icons8.com/) - 提供 PNG 下載
- [IconFinder](https://www.iconfinder.com/) - 大量免費圖示

### 方式 2：使用程式碼生成簡單圖示
可使用 PowerShell + .NET 生成基本圖示：

```powershell
# 此功能需要安裝圖形處理庫
# 或使用 Python PIL/Pillow 生成
```

### 方式 3：使用 Emoji 轉圖片
- ⚙️ → cobie_config (設定)
- 📤 → cobie_export (匯出)
- 📥 → cobie_import (匯入)

## 臨時解決方案

如果暫時沒有圖示檔案，程式會正常運作，只是按鈕沒有圖示顯示。
`IconLoader.TryLoad()` 會回傳 null，不會造成錯誤。

## 安裝圖示後

1. 確保檔案放在 `Resources/Icons/` 資料夾
2. 重新編譯專案
3. 圖示會自動嵌入到 DLL 中
4. 在 Revit 中重新載入外掛即可看到圖示
