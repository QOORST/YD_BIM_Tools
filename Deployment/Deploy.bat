@echo off
REM ====================================
REM YD_RevitTools.LicenseManager 部署腳本
REM 支援 Revit 2022-2026
REM ====================================

echo.
echo ========================================
echo YD BIM 工具 - 授權管理系統部署
echo ========================================
echo.

REM 設定來源路徑（根據建置組態調整）
set "SOURCE_2022=..\bin\Release2022"
set "SOURCE_2023=..\bin\Release2023"
set "SOURCE_2024=..\bin\Release2024"
set "SOURCE_2025=..\bin\Release2025"
set "SOURCE_2026=..\bin\Release2026"

REM 設定目標路徑
set "TARGET_2022=C:\ProgramData\Autodesk\Revit\Addins\2022"
set "TARGET_2023=C:\ProgramData\Autodesk\Revit\Addins\2023"
set "TARGET_2024=C:\ProgramData\Autodesk\Revit\Addins\2024"
set "TARGET_2025=C:\ProgramData\Autodesk\Revit\Addins\2025"
set "TARGET_2026=C:\ProgramData\Autodesk\Revit\Addins\2026"

echo 請選擇要部署的 Revit 版本：
echo.
echo [1] Revit 2022
echo [2] Revit 2023
echo [3] Revit 2024
echo [4] Revit 2025
echo [5] Revit 2026
echo [6] 全部版本
echo [0] 取消
echo.
set /p choice="請輸入選項 (0-6): "

if "%choice%"=="0" goto :END
if "%choice%"=="1" call :DEPLOY 2022 "%SOURCE_2022%" "%TARGET_2022%"
if "%choice%"=="2" call :DEPLOY 2023 "%SOURCE_2023%" "%TARGET_2023%"
if "%choice%"=="3" call :DEPLOY 2024 "%SOURCE_2024%" "%TARGET_2024%"
if "%choice%"=="4" call :DEPLOY 2025 "%SOURCE_2025%" "%TARGET_2025%"
if "%choice%"=="5" call :DEPLOY 2026 "%SOURCE_2026%" "%TARGET_2026%"
if "%choice%"=="6" (
    call :DEPLOY 2022 "%SOURCE_2022%" "%TARGET_2022%"
    call :DEPLOY 2023 "%SOURCE_2023%" "%TARGET_2023%"
    call :DEPLOY 2024 "%SOURCE_2024%" "%TARGET_2024%"
    call :DEPLOY 2025 "%SOURCE_2025%" "%TARGET_2025%"
    call :DEPLOY 2026 "%SOURCE_2026%" "%TARGET_2026%"
)

echo.
echo ========================================
echo 部署完成！
echo ========================================
pause
goto :END

:DEPLOY
set version=%~1
set source=%~2
set target=%~3

echo.
echo ----------------------------------------
echo 正在部署到 Revit %version%...
echo ----------------------------------------

REM 檢查來源檔案
if not exist "%source%\YD_RevitTools.LicenseManager.dll" (
    echo [錯誤] 找不到 DLL 檔案: %source%\YD_RevitTools.LicenseManager.dll
    echo [提示] 請先建置 Release%version% 組態
    goto :EOF
)

REM 檢查目標資料夾
if not exist "%target%" (
    echo [錯誤] Revit %version% Addins 資料夾不存在
    echo [提示] 請確認已安裝 Revit %version%
    goto :EOF
)

REM 複製 DLL
echo 複製 DLL 檔案...
copy /Y "%source%\YD_RevitTools.LicenseManager.dll" "%target%\" > nul
if errorlevel 1 (
    echo [錯誤] 複製 DLL 失敗
    goto :EOF
)

REM 複製 Newtonsoft.Json.dll
if exist "%source%\Newtonsoft.Json.dll" (
    echo 複製 Newtonsoft.Json.dll...
    copy /Y "%source%\Newtonsoft.Json.dll" "%target%\" > nul
)

REM 複製 .addin 檔案
echo 複製 .addin 檔案...
copy /Y "YD_RevitTools.LicenseManager.%version%.addin" "%target%\" > nul
if errorlevel 1 (
    echo [錯誤] 複製 .addin 檔案失敗
    goto :EOF
)

echo [成功] Revit %version% 部署完成！
goto :EOF

:END
exit /b
