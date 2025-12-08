# YD BIM Tools v2.2.4 發布腳本
# 此腳本會自動提交變更並建立 Git 標籤

Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "  YD BIM Tools v2.2.4 發布腳本" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""

# 設定 Git 路徑
$gitPath = "C:\Program Files\Git\bin\git.exe"

# 檢查 Git 是否可用
if (-not (Test-Path $gitPath)) {
    Write-Host "[ERROR] Git 未找到於: $gitPath" -ForegroundColor Red
    Write-Host "請先執行 setup_git.ps1 設定 Git" -ForegroundColor Yellow
    exit 1
}

$gitVersion = & $gitPath --version
Write-Host "[OK] Git 已安裝: $gitVersion" -ForegroundColor Green

Write-Host ""
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "  步驟 1: 檢查變更狀態" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""

& $gitPath status

Write-Host ""
$confirm = Read-Host "是否要提交這些變更？(Y/N)"
if ($confirm -ne "Y" -and $confirm -ne "y") {
    Write-Host "已取消發布" -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "  步驟 2: 提交變更" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""

# 讀取提交訊息
$commitMessage = Get-Content "COMMIT_MESSAGE_v2.2.4.txt" -Raw

# 添加所有變更
& $gitPath add .

# 提交
& $gitPath commit -m $commitMessage

if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] 提交失敗" -ForegroundColor Red
    exit 1
}

Write-Host "[OK] 變更已提交" -ForegroundColor Green

Write-Host ""
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "  步驟 3: 建立 Git 標籤" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""

# 建立標籤
& $gitPath tag -a v2.2.4 -m "Release v2.2.4 - 修復 Revit 2025/2026 COBie 參數建立問題"

if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] 建立標籤失敗" -ForegroundColor Red
    exit 1
}

Write-Host "[OK] 標籤 v2.2.4 已建立" -ForegroundColor Green

Write-Host ""
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "  步驟 4: 推送到 GitHub" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""

$confirmPush = Read-Host "是否要推送到 GitHub？(Y/N)"
if ($confirmPush -ne "Y" -and $confirmPush -ne "y") {
    Write-Host "已跳過推送，您可以稍後手動執行：" -ForegroundColor Yellow
    Write-Host "  & '$gitPath' push origin main" -ForegroundColor Yellow
    Write-Host "  & '$gitPath' push origin v2.2.4" -ForegroundColor Yellow
    exit 0
}

# 推送提交
Write-Host "推送提交到遠端..." -ForegroundColor Cyan
& $gitPath push origin main

if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] 推送提交失敗" -ForegroundColor Red
    exit 1
}

Write-Host "[OK] 提交已推送" -ForegroundColor Green

# 推送標籤
Write-Host "推送標籤到遠端..." -ForegroundColor Cyan
& $gitPath push origin v2.2.4

if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] 推送標籤失敗" -ForegroundColor Red
    exit 1
}

Write-Host "[OK] 標籤已推送" -ForegroundColor Green

Write-Host ""
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "  發布完成！" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "接下來請在 GitHub 上完成以下步驟：" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. 前往 GitHub Releases 頁面" -ForegroundColor White
Write-Host "2. 點擊 'Draft a new release'" -ForegroundColor White
Write-Host "3. 選擇標籤 'v2.2.4'" -ForegroundColor White
Write-Host "4. 複製 RELEASE_NOTES_v2.2.4.md 的內容到描述欄位" -ForegroundColor White
Write-Host "5. 上傳 Output\YD_BIM_Tools_v2.2.4_Setup.exe" -ForegroundColor White
Write-Host "6. 點擊 'Publish release'" -ForegroundColor White
Write-Host ""
Write-Host "詳細步驟請參考：GITHUB_RELEASE_GUIDE.md" -ForegroundColor Cyan
Write-Host ""
Write-Host "安裝程式資訊：" -ForegroundColor Yellow
Write-Host "  檔案：Output\YD_BIM_Tools_v2.2.4_Setup.exe" -ForegroundColor White
Write-Host "  大小：2.88 MB (2,883,340 bytes)" -ForegroundColor White
Write-Host "  MD5： BD0AD4E80EC58155C9004B93488BC8E6" -ForegroundColor White
Write-Host ""

