# ====================================
# YD_RevitTools.LicenseManager 部署腳本 (PowerShell)
# 支援 Revit 2022-2026
# ====================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "YD BIM 工具 - 授權管理系統部署" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 設定來源路徑（根據建置組態調整）
$sourcePaths = @{
    "2022" = "..\bin\Release2022"
    "2023" = "..\bin\Release2023"
    "2024" = "..\bin\Release2024"
    "2025" = "..\bin\Release2025"
    "2026" = "..\bin\Release2026"
}

# 設定目標路徑
$targetPaths = @{
    "2022" = "C:\ProgramData\Autodesk\Revit\Addins\2022"
    "2023" = "C:\ProgramData\Autodesk\Revit\Addins\2023"
    "2024" = "C:\ProgramData\Autodesk\Revit\Addins\2024"
    "2025" = "C:\ProgramData\Autodesk\Revit\Addins\2025"
    "2026" = "C:\ProgramData\Autodesk\Revit\Addins\2026"
}

# 部署函數
function Deploy-RevitVersion {
    param(
        [string]$version,
        [string]$source,
        [string]$target
    )
    
    Write-Host ""
    Write-Host "----------------------------------------" -ForegroundColor Yellow
    Write-Host "正在部署到 Revit $version..." -ForegroundColor Yellow
    Write-Host "----------------------------------------" -ForegroundColor Yellow
    
    # 檢查來源檔案
    $dllPath = Join-Path $source "YD_RevitTools.LicenseManager.dll"
    if (-not (Test-Path $dllPath)) {
        Write-Host "[錯誤] 找不到 DLL 檔案: $dllPath" -ForegroundColor Red
        Write-Host "[提示] 請先建置 Release$version 組態" -ForegroundColor Yellow
        return $false
    }
    
    # 檢查目標資料夾
    if (-not (Test-Path $target)) {
        Write-Host "[錯誤] Revit $version Addins 資料夾不存在" -ForegroundColor Red
        Write-Host "[提示] 請確認已安裝 Revit $version" -ForegroundColor Yellow
        return $false
    }
    
    try {
        # 複製 DLL
        Write-Host "複製 DLL 檔案..." -ForegroundColor Gray
        Copy-Item $dllPath -Destination $target -Force
        
        # 複製 Newtonsoft.Json.dll
        $jsonPath = Join-Path $source "Newtonsoft.Json.dll"
        if (Test-Path $jsonPath) {
            Write-Host "複製 Newtonsoft.Json.dll..." -ForegroundColor Gray
            Copy-Item $jsonPath -Destination $target -Force
        }
        
        # 複製 .addin 檔案
        $addinPath = "YD_RevitTools.LicenseManager.$version.addin"
        if (Test-Path $addinPath) {
            Write-Host "複製 .addin 檔案..." -ForegroundColor Gray
            Copy-Item $addinPath -Destination $target -Force
        }
        
        Write-Host "[成功] Revit $version 部署完成！" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "[錯誤] 部署失敗: $_" -ForegroundColor Red
        return $false
    }
}

# 顯示選單
Write-Host "請選擇要部署的 Revit 版本：" -ForegroundColor White
Write-Host ""
Write-Host "[1] Revit 2022" -ForegroundColor White
Write-Host "[2] Revit 2023" -ForegroundColor White
Write-Host "[3] Revit 2024" -ForegroundColor White
Write-Host "[4] Revit 2025" -ForegroundColor White
Write-Host "[5] Revit 2026" -ForegroundColor White
Write-Host "[6] 全部版本" -ForegroundColor White
Write-Host "[0] 取消" -ForegroundColor White
Write-Host ""

$choice = Read-Host "請輸入選項 (0-6)"

$success = $true

switch ($choice) {
    "0" { 
        Write-Host "已取消" -ForegroundColor Yellow
        exit 
    }
    "1" { 
        $success = Deploy-RevitVersion "2022" $sourcePaths["2022"] $targetPaths["2022"]
    }
    "2" { 
        $success = Deploy-RevitVersion "2023" $sourcePaths["2023"] $targetPaths["2023"]
    }
    "3" { 
        $success = Deploy-RevitVersion "2024" $sourcePaths["2024"] $targetPaths["2024"]
    }
    "4" { 
        $success = Deploy-RevitVersion "2025" $sourcePaths["2025"] $targetPaths["2025"]
    }
    "5" { 
        $success = Deploy-RevitVersion "2026" $sourcePaths["2026"] $targetPaths["2026"]
    }
    "6" {
        foreach ($version in @("2022", "2023", "2024", "2025", "2026")) {
            $result = Deploy-RevitVersion $version $sourcePaths[$version] $targetPaths[$version]
            if (-not $result) { $success = $false }
        }
    }
    default {
        Write-Host "無效的選項" -ForegroundColor Red
        exit
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
if ($success) {
    Write-Host "部署完成！" -ForegroundColor Green
} else {
    Write-Host "部署過程中發生錯誤" -ForegroundColor Red
}
Write-Host "========================================" -ForegroundColor Cyan

Read-Host "按 Enter 鍵退出"
