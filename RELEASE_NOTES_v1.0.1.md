# FluentNPOI v1.0.1 Release Notes

## 🎉 重大更新：專案重命名

本次發布將專案從 **NPOIPlus** 重命名為 **FluentNPOI**，以更好地反映專案的核心特性——流暢（Fluent）風格的 API 設計。

## 📝 變更內容

### ✨ 主要變更

- **專案重命名**：從 `NPOIPlus` 重命名為 `FluentNPOI`
- **命名空間更新**：
  - `NPOIPlusUnitTest` → `FluentNPOIUnitTest`
  - `NPOIPlusConsoleExample` → `FluentNPOIConsoleExample`
- **專案檔案更新**：
  - `NPOIPlus.csproj` → `FluentNPOI.csproj`
  - `NPOIPlusConsoleExample.csproj` → `FluentNPOIConsoleExample.csproj`
  - `NPOIPlusUnitTest.csproj` → `FluentNPOIUnitTest.csproj`
- **GitHub 倉庫 URL 更新**：更新為新的倉庫地址

### 🔧 技術細節

- 所有程式碼中的命名空間已更新
- 專案檔案已正確重命名並更新引用
- 解決方案檔案引用已驗證正確
- 專案建置測試通過 ✅

## 🚀 功能特性（保持不變）

本次更新僅為重命名，所有功能保持不變：

- ✅ 流暢 API 設計
- ✅ 強型別支援（`List<T>` 和 `DataTable`）
- ✅ 樣式緩存機制
- ✅ 動態樣式設定
- ✅ 完整的讀寫功能
- ✅ 多種資料類型支援
- ✅ 公式讀寫支援
- ✅ 豐富的擴展方法

## 📦 NuGet 套件

**套件名稱**：`FluentNPOI`  
**版本**：1.0.1  
**目標框架**：.NET Standard 2.0

## 🔄 遷移指南

如果您之前使用 NPOIPlus，請注意以下變更：

### 命名空間變更

```csharp
// 舊版本
using NPOIPlus;
using NPOIPlus.Models;

// 新版本
using FluentNPOI;
using FluentNPOI.Models;
```

### NuGet 套件

```bash
# 卸載舊套件
Uninstall-Package NPOIPlus

# 安裝新套件
Install-Package FluentNPOI
```

### 程式碼變更

API 使用方式保持完全一致，無需修改業務邏輯代碼：

```csharp
// 使用方式完全相同
var fluent = new FluentWorkbook(new XSSFWorkbook());
fluent.UseSheet("Sheet1")
    .SetCellPosition(ExcelColumns.A, 1)
    .SetValue("Hello FluentNPOI");
```

## 🐛 已知問題

無

## 📚 相關資源

- [GitHub 倉庫](https://github.com/HouseAlwaysWin/FluentNPOI)
- [NuGet 套件](https://www.nuget.org/packages/FluentNPOI)
- [完整文檔](https://github.com/HouseAlwaysWin/FluentNPOI/blob/main/README.md)

## 🙏 感謝

感謝所有使用者和貢獻者的支持！

---

**發布日期**：2024年  
**維護者**：HouseAlwaysWin

