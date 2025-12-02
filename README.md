# FluentNPOI

[![CI](https://github.com/HouseAlwaysWin/NPOIPlus/workflows/CI/badge.svg)](https://github.com/HouseAlwaysWin/NPOIPlus/actions/workflows/ci.yml)
[![.NET Standard 2.0](https://img.shields.io/badge/.NET%20Standard-2.0-blue.svg)](https://docs.microsoft.com/en-us/dotnet/standard/net-standard)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**FluentNPOI** 是基於 [NPOI](https://github.com/dotnetcore/NPOI) 的流暢（Fluent）風格 Excel 操作庫，提供更直觀、更易用的 API 來讀寫 Excel 文件。

[English](#english) | [繁體中文](#繁體中文)

---

## 繁體中文

### 🚀 特性

- ✅ **流暢 API** - 支援鏈式調用，代碼更簡潔易讀
- ✅ **強型別支援** - 完整的泛型支援，支援 `List<T>` 和 `DataTable`
- ✅ **樣式管理** - 強大的樣式緩存機制，避免樣式數量超限
- ✅ **動態樣式** - 支援根據資料動態設置單元格樣式
- ✅ **讀寫功能** - 完整的 Excel 讀取和寫入支援
- ✅ **多種資料類型** - 自動處理字串、數字、日期、布林值等
- ✅ **公式支援** - 支援設置和讀取單元格公式
- ✅ **擴展方法** - 豐富的擴展方法簡化常見操作

### 📦 安裝

```bash
# 使用 NuGet Package Manager
Install-Package FluentNPOI

# 使用 .NET CLI
dotnet add package FluentNPOI
```

### 🎯 快速開始

#### 基本寫入

```csharp
using FluentNPOI;
using NPOI.XSSF.UserModel;
using FluentNPOI.Models;

// 創建 Workbook
var fluent = new FluentWorkbook(new XSSFWorkbook());

// 設置全局樣式
fluent.SetupGlobalCachedCellStyles((workbook, style) =>
{
    style.SetAligment(HorizontalAlignment.Center);
    style.SetBorderAllStyle(BorderStyle.Thin);
});

// 使用工作表並寫入資料
fluent.UseSheet("Sheet1")
    .SetCellPosition(ExcelColumns.A, 1)
    .SetValue("Hello World!");

// 儲存檔案
fluent.SaveToPath("output.xlsx");
```

#### 寫入表格資料

```csharp
var data = new List<Student>
{
    new Student { ID = 1, Name = "Alice", Score = 95.5, IsActive = true },
    new Student { ID = 2, Name = "Bob", Score = 87.0, IsActive = false }
};

fluent.UseSheet("Students")
    .SetTable(data, ExcelColumns.A, 1)

    .BeginTitleSet("學號").SetCellStyle("HeaderStyle")
    .BeginBodySet("ID").SetCellType(CellType.Numeric).End()

    .BeginTitleSet("姓名").SetCellStyle("HeaderStyle")
    .BeginBodySet("Name").End()

    .BeginTitleSet("分數").SetCellStyle("HeaderStyle")
    .BeginBodySet("Score").SetCellType(CellType.Numeric).End()

    .BeginTitleSet("狀態").SetCellStyle("HeaderStyle")
    .BeginBodySet("IsActive").SetCellType(CellType.Boolean).End()

    .BuildRows();
```

#### 讀取 Excel 資料

```csharp
// 開啟現有檔案
var fluent = new FluentWorkbook(new XSSFWorkbook("data.xlsx"));
var sheet = fluent.UseSheet("Sheet1");

// 讀取單一單元格
string name = sheet.GetCellValue<string>(ExcelColumns.A, 1);
int id = sheet.GetCellValue<int>(ExcelColumns.B, 1);
DateTime date = sheet.GetCellValue<DateTime>(ExcelColumns.C, 1);

// 讀取多列資料
for (int row = 2; row <= 10; row++)
{
    var id = sheet.GetCellValue<int>(ExcelColumns.A, row);
    var name = sheet.GetCellValue<string>(ExcelColumns.B, row);
    Console.WriteLine($"ID: {id}, Name: {name}");
}
```

#### 使用 FluentCell 進行鏈式操作

```csharp
fluent.UseSheet("Sheet1")
    .SetCellPosition(ExcelColumns.A, 1)
    .SetValue("Test")
    .SetCellStyle("MyStyle")
    .GetValue<string>(); // 立即讀取剛設置的值
```

### 📚 主要功能

#### 1. 樣式管理

**預定義樣式**

```csharp
fluent.SetupCellStyle("HeaderBlue", (workbook, style) =>
{
    style.SetAligment(HorizontalAlignment.Center);
    style.FillPattern = FillPattern.SolidForeground;
    style.SetCellFillForegroundColor(IndexedColors.LightBlue);
    style.SetBorderAllStyle(BorderStyle.Thin);
});
```

**動態樣式（根據資料變化）**

```csharp
.BeginBodySet("Status")
.SetCellStyle((styleParams) =>
{
    var item = styleParams.GetRowItem<Student>();

    if (item.Score >= 90)
    {
        return new CellStyleConfig("HighScore", style =>
        {
            style.SetCellFillForegroundColor(IndexedColors.LightGreen);
        });
    }
    return new CellStyleConfig("NormalScore", style =>
    {
        style.SetCellFillForegroundColor(IndexedColors.White);
    });
})
.End()
```

#### 2. 資料綁定

**支援 List<T>**

```csharp
List<Employee> employees = GetEmployees();

fluent.UseSheet("Employees")
    .SetTable(employees, ExcelColumns.A, 1)
    .BeginTitleSet("姓名")
    .BeginBodySet("Name").End()
    .BuildRows();
```

**支援 DataTable**

```csharp
DataTable dt = new DataTable();
dt.Columns.Add("ID", typeof(int));
dt.Columns.Add("Name", typeof(string));
dt.Rows.Add(1, "Alice");
dt.Rows.Add(2, "Bob");

fluent.UseSheet("DataTableSheet")
    .SetTable<DataRow>(dt.Rows.Cast<DataRow>(), ExcelColumns.A, 1)
    .BeginTitleSet("編號")
    .BeginBodySet("ID").End()
    .BeginTitleSet("姓名")
    .BeginBodySet("Name").End()
    .BuildRows();
```

#### 3. 單元格操作

**設置值**

```csharp
// 字串
sheet.SetCellPosition(ExcelColumns.A, 1).SetValue("Text");

// 數字
sheet.SetCellPosition(ExcelColumns.B, 1).SetValue(123.45);

// 日期
sheet.SetCellPosition(ExcelColumns.C, 1).SetValue(DateTime.Now);

// 布林值
sheet.SetCellPosition(ExcelColumns.D, 1).SetValue(true);

// 公式
sheet.SetCellPosition(ExcelColumns.E, 1).SetFormulaValue("=A1+B1");
```

**讀取值**

```csharp
// 讀取為特定類型
string text = sheet.GetCellValue<string>(ExcelColumns.A, 1);
double number = sheet.GetCellValue<double>(ExcelColumns.B, 1);
DateTime date = sheet.GetCellValue<DateTime>(ExcelColumns.C, 1);
bool flag = sheet.GetCellValue<bool>(ExcelColumns.D, 1);

// 讀取公式
string formula = sheet.GetCellFormula(ExcelColumns.E, 1);

// 讀取為 object（自動判斷類型）
object value = sheet.GetCellValue(ExcelColumns.A, 1);
```

#### 4. 工作表操作

**設置欄寬**

```csharp
// 單一欄位
sheet.SetColumnWidth(ExcelColumns.A, 20);

// 多個欄位
sheet.SetColumnWidth(ExcelColumns.A, ExcelColumns.E, 15);
```

**合併儲存格**

```csharp
// 橫向合併
sheet.SetExcelCellMerge(ExcelColumns.A, ExcelColumns.C, 1);

// 縱向合併
sheet.SetExcelCellMerge(ExcelColumns.A, ExcelColumns.A, 1, 5);

// 區域合併
sheet.SetExcelCellMerge(ExcelColumns.A, ExcelColumns.C, 1, 3);
```

#### 5. 擴展方法

**顏色設置**

```csharp
style.SetCellFillForegroundColor(255, 0, 0); // RGB
style.SetCellFillForegroundColor("#FF0000"); // Hex
style.SetCellFillForegroundColor(IndexedColors.Red); // 預設顏色
```

**字型設置**

```csharp
style.SetFontInfo(workbook,
    fontFamily: "Arial",
    fontHeight: 12,
    isBold: true,
    color: IndexedColors.Black);
```

**邊框設置**

```csharp
style.SetBorderAllStyle(BorderStyle.Thin); // 所有邊框
style.SetBorderStyle(
    top: BorderStyle.Thick,
    right: BorderStyle.Thin,
    bottom: BorderStyle.Thin,
    left: BorderStyle.Thin
);
```

**對齊設置**

```csharp
style.SetAligment(HorizontalAlignment.Center, VerticalAlignment.Center);
```

**資料格式**

```csharp
style.SetDataFormat(workbook, "yyyy-MM-dd"); // 日期
style.SetDataFormat(workbook, "#,##0.00"); // 數字
```

### 🎨 進階範例

#### 條件格式化

```csharp
fluent.UseSheet("Report")
    .SetTable(salesData, ExcelColumns.A, 1)

    .BeginTitleSet("銷售額")
    .BeginBodySet("Amount")
    .SetCellStyle((styleParams) =>
    {
        var sale = styleParams.GetRowItem<Sale>();

        if (sale.Amount > 10000)
            return new("HighSales", s => s.SetCellFillForegroundColor("#90EE90"));
        else if (sale.Amount > 5000)
            return new("MediumSales", s => s.SetCellFillForegroundColor("#FFFFE0"));
        else
            return new("LowSales", s => s.SetCellFillForegroundColor("#FFB6C1"));
    })
    .End()

    .BuildRows();
```

#### 複製樣式

```csharp
fluent.UseSheet("Sheet2")
    .SetTable(data, ExcelColumns.A, 1)

    // 從 Sheet1 的 A1 複製樣式
    .BeginTitleSet("標題").CopyStyleFromCell(ExcelColumns.A, 1)
    .BeginBodySet("Name").End()

    .BuildRows();
```

#### 多工作表操作

```csharp
var fluent = new FluentWorkbook(new XSSFWorkbook());

// Sheet1
fluent.UseSheet("Summary")
    .SetCellPosition(ExcelColumns.A, 1)
    .SetValue("總覽");

// Sheet2（新建）
fluent.UseSheet("Details", createIfNotExists: true)
    .SetTable(detailData, ExcelColumns.A, 1)
    .BuildRows();

// Sheet3
fluent.UseSheetAt(0) // 使用索引選擇工作表
    .SetCellPosition(ExcelColumns.B, 1)
    .SetValue("Updated");

fluent.SaveToPath("multi-sheet.xlsx");
```

### 📖 API 參考

#### FluentWorkbook

| 方法                                            | 說明                           |
| ----------------------------------------------- | ------------------------------ |
| `UseSheet(string name)`                         | 使用指定名稱的工作表           |
| `UseSheet(string name, bool createIfNotExists)` | 使用工作表，不存在時可選擇創建 |
| `UseSheetAt(int index)`                         | 使用指定索引的工作表           |
| `SetupGlobalCachedCellStyles(Action)`           | 設置全局預設樣式               |
| `SetupCellStyle(string key, Action)`            | 註冊命名樣式                   |
| `GetWorkbook()`                                 | 取得底層 NPOI IWorkbook 物件   |
| `ToStream()`                                    | 輸出為記憶體串流               |
| `SaveToPath(string path)`                       | 儲存到檔案路徑                 |

#### FluentSheet

| 方法                                             | 說明                           |
| ------------------------------------------------ | ------------------------------ |
| `SetCellPosition(ExcelColumns col, int row)`     | 設置當前操作的單元格位置       |
| `GetCellPosition(ExcelColumns col, int row)`     | 取得指定位置的 FluentCell 物件 |
| `GetCellValue<T>(ExcelColumns col, int row)`     | 讀取指定位置的值               |
| `GetCellFormula(ExcelColumns col, int row)`      | 讀取指定位置的公式             |
| `SetTable<T>(IEnumerable<T>, ExcelColumns, int)` | 綁定資料表                     |
| `SetColumnWidth(ExcelColumns col, int width)`    | 設置欄寬                       |
| `SetExcelCellMerge(...)`                         | 合併儲存格                     |
| `GetSheet()`                                     | 取得底層 NPOI ISheet 物件      |

#### FluentCell

| 方法                            | 說明                           |
| ------------------------------- | ------------------------------ |
| `SetValue<T>(T value)`          | 設置單元格值                   |
| `SetFormulaValue(object value)` | 設置公式                       |
| `SetCellStyle(string key)`      | 套用命名樣式                   |
| `SetCellStyle(Func<...>)`       | 套用動態樣式                   |
| `SetCellType(CellType type)`    | 設置單元格類型                 |
| `GetValue()`                    | 讀取單元格值（返回 object）    |
| `GetValue<T>()`                 | 讀取單元格值（轉換為指定類型） |
| `GetFormula()`                  | 讀取公式字串                   |
| `GetCell()`                     | 取得底層 NPOI ICell 物件       |

#### FluentTable

| 方法                                | 說明                 |
| ----------------------------------- | -------------------- |
| `BeginTitleSet(string title)`       | 開始設置表頭         |
| `BeginBodySet(string propertyName)` | 開始設置資料欄位     |
| `BuildRows()`                       | 執行資料綁定並生成列 |

#### FluentTableHeader / FluentTableCell

| 方法                                           | 說明                           |
| ---------------------------------------------- | ------------------------------ |
| `SetValue(object value)`                       | 設置固定值                     |
| `SetValue(Func<...>)`                          | 設置動態值                     |
| `SetFormulaValue(...)`                         | 設置公式                       |
| `SetCellStyle(string key)`                     | 套用命名樣式                   |
| `SetCellStyle(Func<...>)`                      | 套用動態樣式                   |
| `SetCellType(CellType type)`                   | 設置單元格類型                 |
| `CopyStyleFromCell(ExcelColumns col, int row)` | 從其他單元格複製樣式           |
| `End()`                                        | 結束當前設置並返回 FluentTable |

### 🔧 樣式緩存機制

FluentNPOI 實現了智能樣式緩存機制，避免 Excel 檔案樣式數量超過 64000 的限制：

```csharp
// ✅ 使用 Key 緩存樣式（推薦）
.SetCellStyle((styleParams) =>
{
    return new CellStyleConfig("unique-key", style =>
    {
        style.SetCellFillForegroundColor(IndexedColors.Yellow);
    });
})

// ❌ 不使用 Key（每次都創建新樣式）
.SetCellStyle((styleParams) =>
{
    return new CellStyleConfig("", style => // 空 key
    {
        style.SetCellFillForegroundColor(IndexedColors.Yellow);
    });
})
```

### 💡 最佳實踐

1. **使用樣式緩存** - 為常用樣式設定 Key，避免重複創建
2. **全局樣式優先** - 使用 `SetupGlobalCachedCellStyles` 設置基礎樣式
3. **命名樣式** - 使用 `SetupCellStyle` 預先註冊常用樣式
4. **動態樣式需要 Key** - 動態樣式函數中返回有 Key 的 `CellStyleConfig`
5. **釋放資源** - 處理完成後及時釋放 Stream 和 Workbook

### 📝 範例專案

完整範例請參考：

- [FluentNPOIConsoleExample](NPOIPlusConsoleExample/Program.cs) - 控制台範例
- [FluentNPOIUnitTest](NPOIPlusUnitTest/UnitTest1.cs) - 單元測試範例

### 🤝 貢獻

歡迎提交 Issue 和 Pull Request！

### 📄 授權

本專案採用 MIT 授權條款 - 詳見 [LICENSE](LICENSE) 檔案

---

## English

### 🚀 Features

- ✅ **Fluent API** - Chainable method calls for cleaner code
- ✅ **Strong Type Support** - Full generic support for `List<T>` and `DataTable`
- ✅ **Style Management** - Powerful style caching mechanism to avoid Excel's 64k style limit
- ✅ **Dynamic Styling** - Conditional formatting based on cell data
- ✅ **Read & Write** - Complete Excel read and write operations
- ✅ **Multiple Data Types** - Automatic handling of strings, numbers, dates, booleans
- ✅ **Formula Support** - Set and read cell formulas
- ✅ **Extension Methods** - Rich extension methods for common operations

### 📦 Installation

```bash
# Using NuGet Package Manager
Install-Package FluentNPOI

# Using .NET CLI
dotnet add package FluentNPOI
```

### 🎯 Quick Start

#### Basic Write

```csharp
using FluentNPOI;
using NPOI.XSSF.UserModel;
using FluentNPOI.Models;

// Create Workbook
var fluent = new FluentWorkbook(new XSSFWorkbook());

// Setup global style
fluent.SetupGlobalCachedCellStyles((workbook, style) =>
{
    style.SetAligment(HorizontalAlignment.Center);
    style.SetBorderAllStyle(BorderStyle.Thin);
});

// Use sheet and write data
fluent.UseSheet("Sheet1")
    .SetCellPosition(ExcelColumns.A, 1)
    .SetValue("Hello World!");

// Save file
fluent.SaveToPath("output.xlsx");
```

#### Write Table Data

```csharp
var data = new List<Student>
{
    new Student { ID = 1, Name = "Alice", Score = 95.5, IsActive = true },
    new Student { ID = 2, Name = "Bob", Score = 87.0, IsActive = false }
};

fluent.UseSheet("Students")
    .SetTable(data, ExcelColumns.A, 1)

    .BeginTitleSet("ID").SetCellStyle("HeaderStyle")
    .BeginBodySet("ID").SetCellType(CellType.Numeric).End()

    .BeginTitleSet("Name").SetCellStyle("HeaderStyle")
    .BeginBodySet("Name").End()

    .BeginTitleSet("Score").SetCellStyle("HeaderStyle")
    .BeginBodySet("Score").SetCellType(CellType.Numeric).End()

    .BeginTitleSet("Status").SetCellStyle("HeaderStyle")
    .BeginBodySet("IsActive").SetCellType(CellType.Boolean).End()

    .BuildRows();
```

#### Read Excel Data

```csharp
// Open existing file
var fluent = new FluentWorkbook(new XSSFWorkbook("data.xlsx"));
var sheet = fluent.UseSheet("Sheet1");

// Read single cell
string name = sheet.GetCellValue<string>(ExcelColumns.A, 1);
int id = sheet.GetCellValue<int>(ExcelColumns.B, 1);
DateTime date = sheet.GetCellValue<DateTime>(ExcelColumns.C, 1);

// Read multiple rows
for (int row = 2; row <= 10; row++)
{
    var id = sheet.GetCellValue<int>(ExcelColumns.A, row);
    var name = sheet.GetCellValue<string>(ExcelColumns.B, row);
    Console.WriteLine($"ID: {id}, Name: {name}");
}
```

### 📚 Main Features

#### 1. Style Management

**Predefined Styles**

```csharp
fluent.SetupCellStyle("HeaderBlue", (workbook, style) =>
{
    style.SetAligment(HorizontalAlignment.Center);
    style.FillPattern = FillPattern.SolidForeground;
    style.SetCellFillForegroundColor(IndexedColors.LightBlue);
    style.SetBorderAllStyle(BorderStyle.Thin);
});
```

**Dynamic Styles**

```csharp
.BeginBodySet("Status")
.SetCellStyle((styleParams) =>
{
    var item = styleParams.GetRowItem<Student>();

    if (item.Score >= 90)
    {
        return new CellStyleConfig("HighScore", style =>
        {
            style.SetCellFillForegroundColor(IndexedColors.LightGreen);
        });
    }
    return new CellStyleConfig("NormalScore", style =>
    {
        style.SetCellFillForegroundColor(IndexedColors.White);
    });
})
.End()
```

#### 2. Data Binding

**Support List<T>**

```csharp
List<Employee> employees = GetEmployees();

fluent.UseSheet("Employees")
    .SetTable(employees, ExcelColumns.A, 1)
    .BeginTitleSet("Name")
    .BeginBodySet("Name").End()
    .BuildRows();
```

**Support DataTable**

```csharp
DataTable dt = new DataTable();
dt.Columns.Add("ID", typeof(int));
dt.Columns.Add("Name", typeof(string));
dt.Rows.Add(1, "Alice");
dt.Rows.Add(2, "Bob");

fluent.UseSheet("DataTableSheet")
    .SetTable<DataRow>(dt.Rows.Cast<DataRow>(), ExcelColumns.A, 1)
    .BeginTitleSet("ID")
    .BeginBodySet("ID").End()
    .BeginTitleSet("Name")
    .BeginBodySet("Name").End()
    .BuildRows();
```

### 📖 API Reference

See the Chinese section above for detailed API documentation.

### 🤝 Contributing

Issues and Pull Requests are welcome!

### 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 相關連結 / Related Links

- [NPOI](https://github.com/dotnetcore/NPOI) - The underlying library
- [Issues](../../issues) - Report bugs or request features
- [Examples](NPOIPlusConsoleExample/Program.cs) - More code examples
