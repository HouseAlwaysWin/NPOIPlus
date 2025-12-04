# FluentNPOI v1.1.0 Release Notes

## 🎉 新功能發布

本次發布新增了**自動判斷最後一行**功能，讓讀取 Excel 表格數據更加便捷！

## ✨ 主要新功能

### 自動判斷最後一行

現在您可以使用 `GetTable<T>` 方法時，無需手動指定結束行號，系統會自動檢測表格的最後一行。

#### 使用方式

**舊方法（仍然可用）：**
```csharp
// 需要手動指定結束行
var readData = sheet.GetTable<ExampleData>(ExcelColumns.A, 2, 13);
```

**新方法（推薦）：**
```csharp
// 自動判斷最後一行，無需指定結束行
var readData = sheet.GetTable<ExampleData>(ExcelColumns.A, 2);
```

#### 功能特點

- ✅ **智能檢測**：自動從最後一行向上查找，找到第一個有數據的行
- ✅ **空行處理**：自動跳過空白單元格和空行
- ✅ **完全兼容**：不影響現有代碼，原有方法仍然可用
- ✅ **高效準確**：智能算法確保讀取到正確的數據範圍

#### 使用範例

```csharp
// 打開 Excel 文件
var fluent = new FluentWorkbook(new XSSFWorkbook("data.xlsx"));
var sheet = fluent.UseSheet("Sheet1");

// 自動讀取所有數據（從第2行開始，自動判斷最後一行）
var data = sheet.GetTable<Student>(ExcelColumns.A, 2);

Console.WriteLine($"成功讀取 {data.Count} 筆資料");
foreach (var student in data)
{
    Console.WriteLine($"ID: {student.ID}, Name: {student.Name}");
}
```

## 📚 文檔更新

### README 同步

- ✅ README.md 英文版已與中文版完全同步
- ✅ 補充了所有功能說明和 API 參考
- ✅ 添加了完整的範例代碼和最佳實踐

## 🧪 測試覆蓋

本次更新新增了 9 個單元測試用例，確保功能穩定可靠：

- 基本功能測試
- 與手動指定結束行的對比測試
- 空行處理測試
- 中間空行處理測試
- 單行數據測試
- 空工作表測試
- 不同數據類型測試
- 大數據集測試（100行）

## 📦 NuGet 套件

**套件名稱**：`FluentNPOI`  
**版本**：1.1.0  
**目標框架**：.NET Standard 2.0

## 🔄 遷移指南

### 無需任何代碼變更

本次更新完全向後兼容，您無需修改任何現有代碼。如果您想使用新功能，只需：

1. 更新 NuGet 套件到 1.1.0
2. 將 `GetTable<T>(col, startRow, endRow)` 改為 `GetTable<T>(col, startRow)`（可選）

### 升級步驟

```bash
# 更新 NuGet 套件
Update-Package FluentNPOI

# 或使用 .NET CLI
dotnet add package FluentNPOI --version 1.1.0
```

## 🐛 已知問題

無

## 📚 相關資源

- [GitHub 倉庫](https://github.com/HouseAlwaysWin/FluentNPOI)
- [NuGet 套件](https://www.nuget.org/packages/FluentNPOI)
- [完整文檔](https://github.com/HouseAlwaysWin/FluentNPOI/blob/main/README.md)

## 🙏 感謝

感謝所有使用者和貢獻者的支持與反饋！

---

**發布日期**：2025 年 12 月  
**維護者**：HouseAlwaysWin

