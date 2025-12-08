# Git 快速入門指南

✅ **Git 已成功安裝！** (版本 2.52.0)

---

## 📋 快速設定步驟

### 步驟 1：設定 Git 使用者資訊

在 PowerShell 中執行以下命令（請替換成您的資訊）：

```powershell
& "C:\Program Files\Git\bin\git.exe" config --global user.name "Your Name"
& "C:\Program Files\Git\bin\git.exe" config --global user.email "your.email@example.com"
```

### 步驟 2：檢查是否已是 Git 倉庫

```powershell
& "C:\Program Files\Git\bin\git.exe" status
```

如果顯示 "fatal: not a git repository"，表示尚未初始化，請執行：

```powershell
& "C:\Program Files\Git\bin\git.exe" init
```

### 步驟 3：添加遠端 GitHub 倉庫（如果有）

```powershell
& "C:\Program Files\Git\bin\git.exe" remote add origin https://github.com/username/repo.git
```

---

## 🚀 發布 v2.2.4 到 GitHub

### 選項 A：使用自動腳本（推薦）

```powershell
.\publish_v2.2.4.ps1
```

此腳本會自動：
1. 顯示變更狀態
2. 提交所有變更
3. 建立 v2.2.4 標籤
4. 推送到 GitHub

### 選項 B：手動執行命令

```powershell
# 1. 查看變更狀態
& "C:\Program Files\Git\bin\git.exe" status

# 2. 添加所有變更
& "C:\Program Files\Git\bin\git.exe" add .

# 3. 提交變更（使用準備好的提交訊息）
& "C:\Program Files\Git\bin\git.exe" commit -F COMMIT_MESSAGE_v2.2.4.txt

# 4. 建立標籤
& "C:\Program Files\Git\bin\git.exe" tag -a v2.2.4 -m "Release v2.2.4 - Fix Revit 2025/2026 COBie parameter creation"

# 5. 推送提交到 GitHub
& "C:\Program Files\Git\bin\git.exe" push origin main

# 6. 推送標籤到 GitHub
& "C:\Program Files\Git\bin\git.exe" push origin v2.2.4
```

---

## 📦 在 GitHub 上建立 Release

推送完成後，前往 GitHub 網站：

1. 開啟您的專案頁面
2. 點擊 "Releases" → "Draft a new release"
3. 選擇標籤 `v2.2.4`
4. 標題：`YD BIM Tools v2.2.4 - 修復 Revit 2025/2026 COBie 參數建立問題`
5. 描述：複製 `RELEASE_NOTES_v2.2.4.md` 的內容
6. 上傳 `Output\YD_BIM_Tools_v2.2.4_Setup.exe`
7. 點擊 "Publish release"

---

## 💡 常用 Git 命令

```powershell
# 查看狀態
& "C:\Program Files\Git\bin\git.exe" status

# 查看提交歷史
& "C:\Program Files\Git\bin\git.exe" log --oneline

# 查看遠端倉庫
& "C:\Program Files\Git\bin\git.exe" remote -v

# 查看所有標籤
& "C:\Program Files\Git\bin\git.exe" tag

# 查看當前分支
& "C:\Program Files\Git\bin\git.exe" branch
```

---

## ⚠️ 注意事項

1. **首次推送**：如果是第一次推送到 GitHub，可能需要登入 GitHub 帳號
2. **分支名稱**：如果您的主分支不是 `main`，請將命令中的 `main` 替換成您的分支名稱（如 `master`）
3. **遠端名稱**：如果您的遠端倉庫不是 `origin`，請替換成正確的名稱

---

## 📚 詳細文檔

- **完整發布指南**：`GITHUB_RELEASE_GUIDE.md`
- **Release Notes**：`RELEASE_NOTES_v2.2.4.md`
- **發布總結**：`RELEASE_SUMMARY_v2.2.4.md`

---

**準備好了！請按照上述步驟發布 v2.2.4 版本。** 🚀

