# NuGet 發布指南 / NuGet Publishing Guide

本文檔說明如何將 FluentNPOI 發布到 NuGet.org。

This document explains how to publish FluentNPOI to NuGet.org.

---

## 📋 前置準備 / Prerequisites

### 1. 確保項目配置正確

確認 `FluentNPOI/FluentNPOI.csproj` 包含以下 NuGet 包元數據：

- ✅ `PackageId` - 包名稱（必須唯一）
- ✅ `Version` - 版本號
- ✅ `Authors` - 作者
- ✅ `Description` - 包描述
- ✅ `PackageTags` - 標籤
- ✅ `PackageLicenseExpression` - 許可證
- ✅ `PackageProjectUrl` - 項目 URL

### 2. 獲取 NuGet API Key

1. 前往 [NuGet.org](https://www.nuget.org/)
2. 登入您的帳號
3. 點擊右上角頭像 → **Account Settings**
4. 選擇 **API Keys** 標籤
5. 點擊 **Create** 創建新的 API Key
6. 填寫資訊：
   - **Key name**: `FluentNPOI GitHub Actions`（或任何名稱）
   - **Select scopes**: 選擇 **Push new packages and package versions**
   - **Select packages**: 選擇 **All packages**（推薦）
   - **Expires**: 選擇過期時間（建議選擇較長時間，如 1 年）
7. 點擊 **Create**
8. **重要**：複製生成的 API Key（只會顯示一次！）

### 3. 設置 GitHub Secret

1. 前往您的 GitHub 倉庫
2. 點擊 **Settings** → **Secrets and variables** → **Actions**
3. 點擊 **New repository secret**
4. 填寫：
   - **Name**: `NUGET_API_KEY`
   - **Secret**: 貼上剛才複製的 NuGet API Key
5. 點擊 **Add secret**

---

## 🚀 發布方法 / Publishing Methods

### 方法 1：通過 GitHub Release（推薦）⭐

這是最簡單且推薦的方法，會自動觸發發布工作流。

#### 步驟：

1. **確保所有測試通過**

   ```bash
   dotnet test
   ```

2. **更新版本號**

   編輯 `FluentNPOI/FluentNPOI.csproj`，更新 `<Version>` 標籤：

   ```xml
   <Version>1.0.1</Version>
   ```

3. **提交並推送更改**

   ```bash
   git add FluentNPOI/FluentNPOI.csproj
   git commit -m "chore: bump version to 1.0.1"
   git push origin main
   ```

4. **創建 GitHub Release**

   - 前往倉庫的 **Releases** 頁面
   - 點擊 **Draft a new release**
   - 填寫資訊：
     - **Tag version**: `v1.0.1`（**必須以 `v` 開頭**）
     - **Release title**: `Version 1.0.1` 或 `FluentNPOI 1.0.1`
     - **Description**: 填寫更新說明（可參考 CHANGELOG.md）
   - 點擊 **Publish release**

5. **自動發布**

   - GitHub Actions 會自動觸發 `publish.yml` 工作流
   - 工作流會：
     - 構建專案
     - 運行測試
     - 打包 NuGet 包
     - 發布到 NuGet.org

6. **驗證發布**
   - 前往 [NuGet.org](https://www.nuget.org/packages/FluentNPOI)
   - 確認新版本已出現（可能需要等待幾分鐘）

---

### 方法 2：手動觸發工作流

如果不想創建 Release，可以直接手動觸發工作流。

#### 步驟：

1. **更新版本號**（同方法 1 的步驟 2）

2. **提交並推送**（同方法 1 的步驟 3）

3. **手動觸發工作流**

   - 前往倉庫的 **Actions** 標籤頁
   - 選擇 **Publish to NuGet** 工作流
   - 點擊 **Run workflow**
   - 選擇分支（通常是 `main`）
   - 在 **Version** 輸入框中輸入版本號（例如：`1.0.1`）
   - 點擊 **Run workflow**

4. **等待完成**
   - 工作流會自動執行
   - 查看日誌確認發布成功

---

### 方法 3：本地手動發布（測試用）

僅用於測試，不推薦用於正式發布。

#### 步驟：

1. **構建並打包**

   ```bash
   cd FluentNPOI
   dotnet pack --configuration Release
   ```

2. **發布到 NuGet.org**
   ```bash
   dotnet nuget push bin/Release/FluentNPOI.*.nupkg --api-key YOUR_API_KEY --source https://api.nuget.org/v3/index.json
   ```

---

## 📝 版本號規範 / Version Numbering

遵循 [語義化版本控制 (SemVer)](https://semver.org/lang/zh-CN/)：

- **主版本號 (Major)**: 不兼容的 API 修改
- **次版本號 (Minor)**: 向下兼容的功能性新增
- **修訂號 (Patch)**: 向下兼容的問題修正

範例：

- `1.0.0` - 初始版本
- `1.0.1` - 修復 bug
- `1.1.0` - 新增功能
- `2.0.0` - 重大變更（可能不兼容）

### 預發布版本

如果需要發布預發布版本（如 alpha、beta、rc）：

- `1.0.0-alpha.1`
- `1.0.0-beta.1`
- `1.0.0-rc.1`

在 GitHub Release 中，Tag 仍使用 `v` 前綴：

- Tag: `v1.0.0-alpha.1`
- 工作流會自動去除 `v` 前綴

---

## ✅ 發布檢查清單 / Publishing Checklist

發布前請確認：

- [ ] 所有測試通過
- [ ] 版本號已更新
- [ ] CHANGELOG.md 已更新
- [ ] README.md 已更新（如有需要）
- [ ] 代碼已提交並推送到 GitHub
- [ ] NUGET_API_KEY Secret 已設置
- [ ] 版本號符合 SemVer 規範
- [ ] 包元數據（描述、標籤等）正確

---

## 🔍 驗證發布 / Verifying Publication

### 檢查 NuGet.org

1. 前往 [NuGet.org](https://www.nuget.org/)
2. 搜索 `FluentNPOI`
3. 確認新版本已出現
4. 檢查包資訊是否正確

### 測試安裝

```bash
# 使用 .NET CLI
dotnet add package FluentNPOI --version 1.0.1

# 或使用 Package Manager
Install-Package FluentNPOI -Version 1.0.1
```

---

## 🐛 常見問題 / Troubleshooting

### Q: 發布失敗，提示 "Package already exists"

**A:** 該版本號已存在。解決方法：

- 使用新的版本號
- 或刪除 NuGet.org 上的舊版本（如果允許）

### Q: 發布失敗，提示 "API Key invalid"

**A:** 檢查：

1. GitHub Secret `NUGET_API_KEY` 是否正確設置
2. API Key 是否已過期
3. API Key 權限是否正確（需要 Push 權限）

### Q: 發布成功但 NuGet.org 上看不到

**A:** NuGet.org 需要幾分鐘來索引新包，請稍候再檢查。

### Q: 如何更新包描述或標籤？

**A:** 修改項目文件中的元數據，然後發布新版本。

### Q: 可以撤回已發布的版本嗎？

**A:** NuGet.org 不允許刪除已發布的版本，但可以：

- 發布新版本修復問題
- 將舊版本標記為已棄用（deprecated）

---

## 📚 相關資源 / Resources

- [NuGet 文檔](https://docs.microsoft.com/en-us/nuget/)
- [語義化版本控制](https://semver.org/)
- [GitHub Actions 文檔](https://docs.github.com/en/actions)
- [.NET CLI 文檔](https://docs.microsoft.com/en-us/dotnet/core/tools/)

---

## 🎉 完成！

發布完成後，您的包就可以被全世界的開發者使用了！

After publishing, your package will be available for developers worldwide!

---

**最後更新 / Last Updated:** 2024-12-01

**維護者 / Maintainer:** HouseAlwaysWin
