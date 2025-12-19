using System.Data;
using System;
using System.IO;
using FluentNPOI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using FluentNPOI.Models;
using System.Linq;
using FluentNPOI.Stages;
using FluentNPOI.Streaming.Mapping;

namespace FluentNPOIConsoleExample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var testData = GetTestData();
                var filePath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Test.xlsx";
                var outputPath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Test2_v2.xlsx";

                // Ensure Resources folder exists
                var dir = Path.GetDirectoryName(outputPath);
                if (dir != null) Directory.CreateDirectory(dir);

                var workbook = new XSSFWorkbook();
                var fluent = new FluentWorkbook(workbook);

                // Setup styles
                SetupStyles(fluent);

                // ========== Write Examples ==========
                // Table write examples
                CreateBasicTableExample(fluent, testData);
                CreateSummaryExample(fluent, testData);
                CreateDataTableExample(fluent);

                // Style write examples
                CreateCopyStyleExample(fluent, testData);
                CreateCellStyleRangeExample(fluent);
                CreateSheetGlobalStyleExample(fluent, testData);
                CreateMappingStylingExample(fluent, testData);

                // Cell write examples
                CreateSetCellValueExample(fluent);
                CreateCellMergeExample(fluent);
                CreatePictureExample(fluent);

                // Smart Pipeline example
                CreateSmartPipelineExample(testData);

                // DOM Edit example
                CreateDomEditExample();

                // HTML Export example
                CreateHtmlExportExample(fluent);

                // PDF Export example
                CreatePdfExportExample(fluent);

                // Save file
                fluent.SaveToPath(outputPath);
                Console.WriteLine($"✓ 檔案儲存至: {outputPath}");

                // Read examples
                ReadExcelExamples(fluent);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        #region Test Data

        static List<ExampleData> GetTestData()
        {
            return new List<ExampleData>
            {
                new(1, "Alice Chen", new DateTime(1994, 1, 1), true, 95.5, 12500.75m, "優秀學生"),
                new(2, "Bob Lee", new DateTime(1989, 5, 12), false, 78.0, 8900.50m, "需改進"),
                new(3, "Søren", new DateTime(1985, 7, 23), true, 88.5, 15000.00m, "表現良好"),
                new(4, "王小明", new DateTime(2000, 2, 29), true, 92.0, 11200.80m, "進步快速"),
                new(5, "This is a very very very long name to test wrapping and width", new DateTime(1980, 12, 31), false, 65.5, 7500.25m, "LongName"),
                new(6, "Élodie", new DateTime(1995, 5, 15), true, 85.0, 9800.00m, "穩定發揮"),
                new(7, "O'Connor", new DateTime(1975, 7, 7), false, 72.5, 8200.50m, "待觀察"),
                new(8, "李雷", new DateTime(2010, 10, 10), true, 90.0, 10500.75m, "潛力股"),
                new(9, "山田太郎", new DateTime(1999, 3, 3), true, 87.5, 9500.00m, "穩健型"),
                new(10, "Мария", new DateTime(1988, 8, 8), false, 70.0, 8000.25m, "需加強"),
                new(11, "محمد", new DateTime(1991, 9, 9), true, 93.5, 12000.00m, "頂尖"),
                new(12, "김민준", new DateTime(2004, 4, 4), true, 89.0, 10200.50m, "均衡發展"),
            };
        }

        #endregion

        #region Style Setup

        static void SetupStyles(FluentWorkbook fluent)
        {
            fluent
                // 設定全域基礎樣式
                .SetupGlobalCachedCellStyles((workbook, style) =>
                {
                    style.SetAlignment(HorizontalAlignment.Center);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetFontInfo(workbook, "Calibri", 10);
                })

                // 使用 inheritFrom 繼承 global，只覆寫需要改的屬性
                .SetupCellStyle("BodyString", (workbook, style) =>
                {
                    style.SetFontInfo(workbook, "新細明體", 10);
                }, inheritFrom: "global")

                .SetupCellStyle("DateOfBirth", (workbook, style) =>
                {
                    style.SetDataFormat(workbook, "yyyy-MM-dd");
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                }, inheritFrom: "global")

                .SetupCellStyle("HeaderBlue", (workbook, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightCornflowerBlue);
                }, inheritFrom: "global")

                .SetupCellStyle("BodyGreen", (workbook, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                }, inheritFrom: "global")

                .SetupCellStyle("AmountCurrency", (workbook, style) =>
                {
                    style.SetDataFormat(workbook, "#,##0.00");
                    style.SetAlignment(HorizontalAlignment.Right);
                }, inheritFrom: "global")

                .SetupCellStyle("HighlightYellow", (workbook, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Yellow);
                }, inheritFrom: "global");
        }


        #endregion

        #region Table Write Examples

        /// <summary>
        /// Example 1: Basic table - Demonstrates various field types using FluentMapping
        /// </summary>
        static void CreateBasicTableExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("建立 BasicTableExample (BasicTable)...");

            var mapping = new FluentMapping<ExampleData>();
            mapping.Map(x => x.ID).ToColumn(ExcelCol.A)
                .WithTitle("ID").WithTitleStyle("HeaderBlue").WithStyle("BodyGreen")
                .WithCellType(CellType.Numeric);
            mapping.Map(x => x.Name).ToColumn(ExcelCol.B)
                .WithTitle("名稱").WithTitleStyle("HeaderBlue").WithStyle("BodyGreen");
            mapping.Map(x => x.DateOfBirth).ToColumn(ExcelCol.C)
                .WithTitle("生日").WithTitleStyle("HeaderBlue").WithStyle("DateOfBirth");
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.D)
                .WithTitle("是否活躍").WithTitleStyle("HeaderBlue")
                .WithCellType(CellType.Boolean);
            mapping.Map(x => x.Score).ToColumn(ExcelCol.E)
                .WithTitle("分數").WithTitleStyle("HeaderBlue")
                .WithCellType(CellType.Numeric);
            mapping.Map(x => x.Amount).ToColumn(ExcelCol.F)
                .WithTitle("金額").WithTitleStyle("HeaderBlue").WithStyle("AmountCurrency")
                .WithCellType(CellType.Numeric);
            mapping.Map(x => x.Notes).ToColumn(ExcelCol.G)
                .WithTitle("備註").WithTitleStyle("HeaderBlue");

            mapping.Map(x => x.MaybeNull).ToColumn(ExcelCol.H)
                .WithTitle("可能為空").WithTitleStyle("HeaderBlue");

            fluent.UseSheet("BasicTableExample", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.H, 20)
                .SetTable(testData, mapping)
                .BuildRows();

            Console.WriteLine("  ✓ BasicTableExample 建立完成");
        }

        /// <summary>
        /// Example 2: Summary table
        /// </summary>
        static void CreateSummaryExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("建立 Summary...");

            var mapping = new FluentMapping<ExampleData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A)
                .WithTitle("姓名").WithTitleStyle("HeaderBlue");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B)
                .WithTitle("分數").WithTitleStyle("HeaderBlue").WithStyle("AmountCurrency")
                .WithCellType(CellType.Numeric);
            mapping.Map(x => x.DateOfBirth).ToColumn(ExcelCol.C)
                .WithTitle("日期").WithTitleStyle("HeaderBlue").WithStyle("DateOfBirth");
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.D)
                .WithTitle("狀態").WithTitleStyle("HeaderBlue")
                .WithCellType(CellType.Boolean);
            mapping.Map(x => x.Notes).ToColumn(ExcelCol.E)
                .WithTitle("備註").WithTitleStyle("HeaderBlue").WithStyle("HighlightYellow");

            fluent.UseSheet("Summary", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.E, 20)
                .SetTable(testData, mapping)
                .BuildRows();

            Console.WriteLine("  ✓ Summary 建立完成");
        }

        /// <summary>
        /// Example 3: Using DataTable as data source
        /// </summary>
        static void CreateDataTableExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 DataTableExample...");

            DataTable dataTable = new DataTable("StudentData");
            dataTable.Columns.Add("StudentID", typeof(int));
            dataTable.Columns.Add("StudentName", typeof(string));
            dataTable.Columns.Add("BirthDate", typeof(DateTime));
            dataTable.Columns.Add("IsEnrolled", typeof(bool));
            dataTable.Columns.Add("GPA", typeof(double));
            dataTable.Columns.Add("Tuition", typeof(decimal));
            dataTable.Columns.Add("Department", typeof(string));

            dataTable.Rows.Add(101, "張三", new DateTime(1998, 3, 15), true, 3.8, 25000m, "資訊工程");
            dataTable.Rows.Add(102, "李四", new DateTime(1999, 7, 20), true, 3.5, 22000m, "電機工程");
            dataTable.Rows.Add(103, "王五", new DateTime(1997, 11, 5), false, 2.9, 20000m, "機械工程");
            dataTable.Rows.Add(104, "趙六", new DateTime(2000, 1, 30), true, 3.9, 28000m, "資訊工程");
            dataTable.Rows.Add(105, "陳七", new DateTime(1998, 9, 12), true, 3.6, 23000m, "企業管理");
            dataTable.Rows.Add(106, "林八", new DateTime(1999, 5, 8), false, 3.2, 21000m, "財務金融");

            // 使用 DataTableMapping
            var mapping = new DataTableMapping();
            mapping.Map("StudentID").ToColumn(ExcelCol.A).WithTitle("學號")
                .WithTitleStyle("HeaderBlue").WithCellType(CellType.Numeric);
            mapping.Map("StudentName").ToColumn(ExcelCol.B).WithTitle("姓名")
                .WithTitleStyle("HeaderBlue").WithCellType(CellType.String);
            mapping.Map("BirthDate").ToColumn(ExcelCol.C).WithTitle("出生日期")
                .WithTitleStyle("HeaderBlue").WithStyle("DateOfBirth");
            mapping.Map("IsEnrolled").ToColumn(ExcelCol.D).WithTitle("在學中")
                .WithTitleStyle("HeaderBlue").WithCellType(CellType.Boolean);
            mapping.Map("GPA").ToColumn(ExcelCol.E).WithTitle("GPA")
                .WithTitleStyle("HeaderBlue").WithCellType(CellType.Numeric);
            mapping.Map("Tuition").ToColumn(ExcelCol.F).WithTitle("學費")
                .WithTitleStyle("HeaderBlue").WithStyle("AmountCurrency").WithCellType(CellType.Numeric);
            mapping.Map("Department").ToColumn(ExcelCol.G).WithTitle("科系")
                .WithTitleStyle("HeaderBlue")
                .WithValue((row, excelRow, col) => $"{row["StudentID"]}{excelRow}{col}{row["Department"]} hello")
                .WithStyle("BodyString")
                .WithCellType(CellType.String);

            fluent.UseSheet("DataTableExample", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.G, 20)
                .WriteDataTable(dataTable, mapping);

            Console.WriteLine("  ✓ DataTableExample 建立完成");
        }

        #endregion

        #region Style Write Examples

        /// <summary>
        /// Example 4: Batch set cell range styles
        /// </summary>
        static void CreateCellStyleRangeExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 CellStyleRangeDemo...");

            fluent.UseSheet("CellStyleRangeDemo", true)
                .SetCellStyleRange(new CellStyleConfig("HighlightRed", style =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Red);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                }), ExcelCol.A, ExcelCol.D, 1, 3)
                .SetCellStyleRange(new CellStyleConfig("HighlightOrange", style =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Orange);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                }), ExcelCol.A, ExcelCol.D, 4, 6)
                .SetCellStyleRange(new CellStyleConfig("HighlightYellow", style =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Yellow);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                }), ExcelCol.A, ExcelCol.D, 7, 9)
                .SetCellStyleRange(new CellStyleConfig("HighlightGreen", style =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Green);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                }), ExcelCol.A, ExcelCol.D, 10, 12);

            Console.WriteLine("  ✓ CellStyleRangeDemo 建立完成");
        }

        /// <summary>
        /// Example 3.5: Copy style and dynamic style
        /// </summary>
        static void CreateCopyStyleExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("建立 CopyStyleExample...");

            // 定義幾個動態樣式
            fluent.SetupCellStyle("ActiveGreen", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                style.SetBorderAllStyle(BorderStyle.Thin);
                style.SetAlignment(HorizontalAlignment.Center);
            })
            .SetupCellStyle("InactiveRed", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.SetCellFillForegroundColor(IndexedColors.Rose);
                style.SetBorderAllStyle(BorderStyle.Thin);
                style.SetAlignment(HorizontalAlignment.Center);
            });

            var mapping = new FluentMapping<ExampleData>();

            // 1. 從 Sheet1 的 A1 複製標題樣式 (HeaderBlue)
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A)
                .WithTitle("姓名 (Copy Header)")
                .WithRowOffset(2)
                .WithTitleStyleFrom(1, ExcelCol.B); // Copy from Sheet1!B1 if current sheet is same... wait.
                                                    // WithTitleStyleFrom copies from CURRENT sheet's cell.
                                                    // UseSheet switches sheet. 
                                                    // So we need to put some styles in the current sheet first or copy from self.
                                                    // Let's demo dynamic style first.

            // 調整: 直接用樣式 Key 演示 dynamic style
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.B)
                .WithTitle("狀態 (Dynamic)")
                .WithTitleStyle("HeaderBlue")
                .WithDynamicStyle(item => item.IsActive ? "ActiveGreen" : "InactiveRed");

            // 演示 CopyStyleFromCell: 先在目前 Sheet 建立一個樣板儲存格
            var sheet = fluent.UseSheet("CopyStyleExample", true);
            sheet.SetCellPosition(ExcelCol.Z, 1).SetValue("Template").SetCellStyle("HeaderBlue");

            // 繼續 mapping
            // 注意: 這裡 mapping 定義的 WithTitleStyleFrom 會在 Apply 時去抓
            mapping.Map(x => x.Score).ToColumn(ExcelCol.C)
                .WithTitle("分數 (Copy Z1)")
                .WithTitleStyleFrom(1, ExcelCol.Z) // 從 Z1 複製樣式
                .WithCellType(CellType.Numeric);

            // 使用 WithStartRow 設定預設起始列（而非在 SetTable 傳入）
            mapping.WithStartRow(2);

            sheet.SetColumnWidth(ExcelCol.A, ExcelCol.C, 20)
                .SetTable(testData, mapping) // 會使用 mapping.StartRow (= 2)
                .BuildRows();

            Console.WriteLine("  ✓ CopyStyleExample 建立完成");
        }

        /// <summary>
        /// Example 5: Per-sheet global styles
        /// </summary>
        static void CreateSheetGlobalStyleExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("建立 SheetGlobalStyle...");

            var limitedData = testData.Take(5).ToList();

            var mapping = new FluentMapping<ExampleData>();
            mapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID");
            mapping.Map(x => x.Name).ToColumn(ExcelCol.B).WithTitle("名稱");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.C).WithTitle("分數")
                .WithCellType(CellType.Numeric);
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.D).WithTitle("是否活躍")
                .WithCellType(CellType.Boolean);

            // Sheet 1: 綠色
            fluent.UseSheet("SheetGlobalStyle_Green", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.D, 20)
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetAlignment(HorizontalAlignment.Center);
                })
                .SetTable(limitedData, mapping, 2)
                .BuildRows();

            fluent.UseSheet("SheetGlobalStyle_Green")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Sheet 全域樣式：綠色")
                .SetCellStyle("HeaderBlue");
            fluent.UseSheet("SheetGlobalStyle_Green")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.D, 1);

            // Sheet 2: 黃色
            fluent.UseSheet("SheetGlobalStyle_Yellow", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.D, 20)
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightYellow);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetAlignment(HorizontalAlignment.Center);
                })
                .SetTable(limitedData, mapping, 2)
                .BuildRows();

            fluent.UseSheet("SheetGlobalStyle_Yellow")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Sheet 全域樣式：黃色")
                .SetCellStyle("HeaderBlue");
            fluent.UseSheet("SheetGlobalStyle_Yellow")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.D, 1);

            // Sheet 3: 混合使用 Sheet 全域樣式和特定樣式
            var mixedMapping = new FluentMapping<ExampleData>();
            mixedMapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID")
                .WithStyle("HighlightYellow"); // 覆蓋全域

            mixedMapping.Map(x => x.Name).ToColumn(ExcelCol.B).WithTitle("名稱"); // 使用全域

            mixedMapping.Map(x => x.DateOfBirth).ToColumn(ExcelCol.C).WithTitle("生日")
                .WithStyle("DateOfBirth"); // 覆蓋全域

            mixedMapping.Map(x => x.IsActive).ToColumn(ExcelCol.D).WithTitle("是否活躍")
                .WithCellType(CellType.Boolean)
                .WithDynamicStyle(x => x.IsActive ? "Sheet3_ActiveGreen" : "Sheet3_InactiveRed");

            // 預先註冊 Dynamic Styles
            fluent.SetupCellStyle("Sheet3_ActiveGreen", (wb, s) =>
            {
                s.FillPattern = FillPattern.SolidForeground;
                s.SetCellFillForegroundColor(IndexedColors.Green);
                s.SetBorderAllStyle(BorderStyle.Thin);
                s.SetFontInfo(wb, fontFamily: "新細明體");
            });
            fluent.SetupCellStyle("Sheet3_InactiveRed", (wb, s) =>
            {
                s.FillPattern = FillPattern.SolidForeground;
                s.SetCellFillForegroundColor(IndexedColors.Red);
                s.SetBorderAllStyle(BorderStyle.Thin);
                s.SetFontInfo(wb, fontFamily: "新細明體");
            });

            fluent.UseSheet("SheetGlobalStyle_Mixed", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.D, 20)
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Aqua);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetAlignment(HorizontalAlignment.Center);
                    style.SetFontInfo(wb, fontFamily: "新細明體");
                })
                .SetTable(limitedData, mixedMapping, 2)
                .BuildRows();

            fluent.UseSheet("SheetGlobalStyle_Mixed")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("混合樣式：Sheet 全域(水藍) + 特定樣式覆蓋")
                .SetCellStyle("HeaderBlue");
            fluent.UseSheet("SheetGlobalStyle_Mixed")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.D, 1);

            Console.WriteLine("  ✓ SheetGlobalStyle 建立完成");
        }

        /// <summary>
        /// Example 9: Direct styling in mapping (New Feature)
        /// </summary>
        static void CreateMappingStylingExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("建立 MappingStylingExample...");

            var mapping = new FluentMapping<ExampleData>();

            mapping.Map(x => x.Name)
                .ToColumn(ExcelCol.A)
                .WithTitle("姓名")
                .WithBackgroundColor(IndexedColors.LightCornflowerBlue) // Direct style!
                .WithAlignment(HorizontalAlignment.Center);

            mapping.Map(x => x.Score)
                .ToColumn(ExcelCol.B)
                .WithTitle("分數")
                .WithNumberFormat("0.0") // Direct number format!
                .WithFont(isBold: true);

            mapping.Map(x => x.DateOfBirth)
                .ToColumn(ExcelCol.C)
                .WithTitle("日期")
                .WithNumberFormat("yyyy-mm-dd")
                .WithBackgroundColor(IndexedColors.LightYellow);

            fluent.UseSheet("MappingStylingExample", true)
               .SetColumnWidth(ExcelCol.A, ExcelCol.C, 20)
               .SetTable(testData.Take(5), mapping)
               .BuildRows()
               .SetAutoFilter();

            Console.WriteLine("  ✓ MappingStylingExample 建立完成");
        }

        #endregion

        #region Cell Write Examples

        /// <summary>
        /// Example 6: Set single cell value
        /// </summary>
        static void CreateSetCellValueExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 SetCellValueExample...");

            fluent.UseSheet("SetCellValueExample", true)
                .SetColumnWidth(ExcelCol.A, 20)
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Hello, World!")
                .SetCellStyle("HighlightYellow");

            Console.WriteLine("  ✓ SetCellValueExample 建立完成");
        }

        /// <summary>
        /// Example 7: Cell merge
        /// </summary>
        static void CreateCellMergeExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 CellMergeExample...");

            var sheet = fluent.UseSheet("CellMergeExample", true);
            sheet.SetColumnWidth(ExcelCol.A, ExcelCol.E, 15);

            // 水平合併
            sheet.SetCellPosition(ExcelCol.A, 1)
                .SetValue("銷售報表")
                .SetCellStyle("HeaderBlue");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.E, 1);

            // 設定子標題
            sheet.SetCellPosition(ExcelCol.A, 2).SetValue("產品名稱");
            sheet.SetCellPosition(ExcelCol.B, 2).SetValue("銷售量");
            sheet.SetCellPosition(ExcelCol.C, 2).SetValue("單價");
            sheet.SetCellPosition(ExcelCol.D, 2).SetValue("總金額");
            sheet.SetCellPosition(ExcelCol.E, 2).SetValue("備註");

            for (ExcelCol col = ExcelCol.A; col <= ExcelCol.E; col++)
            {
                sheet.SetCellPosition(col, 2).SetCellStyle("HeaderBlue");
            }

            // 垂直合併
            sheet.SetCellPosition(ExcelCol.A, 3).SetValue("電子產品");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 3, 5);

            sheet.SetCellPosition(ExcelCol.B, 3).SetValue(100);
            sheet.SetCellPosition(ExcelCol.C, 3).SetValue(5000);
            sheet.SetCellPosition(ExcelCol.D, 3).SetValue(500000);
            sheet.SetCellPosition(ExcelCol.E, 3).SetValue("熱銷");

            sheet.SetCellPosition(ExcelCol.B, 4).SetValue(80);
            sheet.SetCellPosition(ExcelCol.C, 4).SetValue(3000);
            sheet.SetCellPosition(ExcelCol.D, 4).SetValue(240000);
            sheet.SetCellPosition(ExcelCol.E, 4).SetValue("穩定");

            sheet.SetCellPosition(ExcelCol.B, 5).SetValue(50);
            sheet.SetCellPosition(ExcelCol.C, 5).SetValue(2000);
            sheet.SetCellPosition(ExcelCol.D, 5).SetValue(100000);
            sheet.SetCellPosition(ExcelCol.E, 5).SetValue("一般");

            // 區域合併
            sheet.SetCellPosition(ExcelCol.A, 6).SetValue("總計");
            sheet.SetCellPosition(ExcelCol.D, 6).SetValue(840000);
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 6);
            sheet.SetCellPosition(ExcelCol.A, 6).SetCellStyle("HighlightYellow");

            // 5. Multiple merge regions example
            sheet.SetCellPosition(ExcelCol.A, 8).SetValue("部門A");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 8, 10); // Vertical merge A8-A10

            sheet.SetCellPosition(ExcelCol.B, 8).SetValue("部門B");
            sheet.SetExcelCellMerge(ExcelCol.B, ExcelCol.B, 8, 10); // Vertical merge B8-B10

            sheet.SetCellPosition(ExcelCol.C, 8).SetValue("部門C");
            sheet.SetExcelCellMerge(ExcelCol.C, ExcelCol.C, 8, 10); // Vertical merge C8-C10

            // 6. Region merge example (multiple rows and columns)
            sheet.SetCellPosition(ExcelCol.A, 12).SetValue("重要通知");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.E, 12, 14); // Merge A12-E14 (region)
            sheet.SetCellPosition(ExcelCol.A, 12).SetCellStyle("HighlightYellow");

            Console.WriteLine("  ✓ CellMergeExample 建立完成");
        }

        /// <summary>
        /// Example 8: Insert picture example
        /// </summary>
        static void CreatePictureExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 PictureExample...");

            var sheet = fluent.UseSheet("PictureExample", true);

            // Set column width
            sheet.SetColumnWidth(ExcelCol.A, ExcelCol.D, 20);

            // Read image file
            var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "pain.jpg");

            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Warning: Image file not found: {imagePath}");
                return;
            }

            byte[] imageBytes = File.ReadAllBytes(imagePath);

            // 1. Basic picture insertion (auto-calculate height, use default column width ratio)
            sheet.SetCellPosition(ExcelCol.A, 1)
                .SetValue("基本插入（自動高度）")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.A, 2)
                .SetPictureOnCell(imageBytes, 200); // Width 200 pixels, height auto-calculated (1:1)

            // 2. Manually set width and height
            sheet.SetCellPosition(ExcelCol.B, 1)
                .SetValue("手動設置尺寸")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.B, 2)
                .SetPictureOnCell(imageBytes, 200, 150); // Width 200, height 150 pixels

            // 3. Custom column width conversion ratio
            sheet.SetCellPosition(ExcelCol.C, 1)
                .SetValue("自定義列寬比例")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.C, 2)
                .SetPictureOnCell(imageBytes, 300, AnchorType.MoveAndResize, 5.0); // Use 5.0 as conversion ratio

            // 4. Chained call example (continue setting after inserting picture)
            sheet.SetCellPosition(ExcelCol.D, 1)
                .SetValue("鏈式調用示例")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.D, 2)
                .SetPictureOnCell(imageBytes, 180, 180, AnchorType.MoveAndResize, 7.0)
                .SetCellPosition(ExcelCol.D, 10)
                .SetValue("圖片下方文字"); // Chained call, set text after picture

            // 5. Different anchor type examples
            sheet.SetCellPosition(ExcelCol.A, 5)
                .SetValue("MoveAndResize（默認）")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.A, 6)
                .SetPictureOnCell(imageBytes, 150, 150, AnchorType.MoveAndResize);

            sheet.SetCellPosition(ExcelCol.B, 5)
                .SetValue("MoveDontResize")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.B, 6)
                .SetPictureOnCell(imageBytes, 150, 150, AnchorType.MoveDontResize);

            sheet.SetCellPosition(ExcelCol.C, 5)
                .SetValue("DontMoveAndResize")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.C, 6)
                .SetPictureOnCell(imageBytes, 150, 150, AnchorType.DontMoveAndResize);

            // 6. Multiple pictures arrangement example
            sheet.SetCellPosition(ExcelCol.A, 9)
                .SetValue("多圖片排列")
                .SetCellStyle("HeaderBlue");

            for (int i = 0; i < 3; i++)
            {
                sheet.SetCellPosition((ExcelCol)((int)ExcelCol.A + i), 10)
                    .SetPictureOnCell(imageBytes, 100, 100);
            }

            Console.WriteLine("  ✓ PictureExample 建立完成");
        }

        /// <summary>
        /// Example 10: Smart Pipeline (Streaming & Legacy)
        /// </summary>
        static void CreateSmartPipelineExample(List<ExampleData> testData)
        {
            Console.WriteLine("建立 SmartPipelineExample...");

            var mapping = new FluentMapping<ExampleData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("姓名");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B).WithTitle("分數");
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.C).WithTitle("狀態");

            // 1. 產生來源檔案 (模擬用)
            var sourceFile = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Source.xlsx";
            var wb = new FluentWorkbook(new XSSFWorkbook());
            wb.UseSheet("Data").SetTable(testData, mapping).BuildRows();
            wb.SaveToFile(sourceFile);

            // 2. 串流處理：輸出為 .xlsx (SXSSF - 高速)
            var outFileXlsx = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Pipeline_Out.xlsx";

            FluentWorkbook.Stream<ExampleData>(sourceFile)
                .Transform(d =>
                {
                    d.Name += " (Streamed)";
                    d.Score += 1.1; // 加分 10%
                })
                .WithMapping(mapping)
                .Configure(sheet =>
                {
                    sheet.SetColumnWidth(ExcelCol.A, 40);
                    sheet.FreezeTitleRow();
                })
                .SaveAs(outFileXlsx);

            Console.WriteLine($"  ✓ Pipeline (XLSX) 處理完成: {outFileXlsx}");

            // 3. 相容處理：輸出為 .xls (HSSF - DOM)
            var outFileXls = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Pipeline_Out.xls";

            FluentWorkbook.Stream<ExampleData>(sourceFile)
                .Transform(d => d.Name += " (Legacy)")
                .WithMapping(mapping)
                .SaveAs(outFileXls);

            Console.WriteLine($"  ✓ Pipeline (XLS) 處理完成: {outFileXls}");
        }

        /// <summary>
        /// Example 11: DOM Edit (Modify existing file)
        /// </summary>
        static void CreateDomEditExample()
        {
            Console.WriteLine("建立 DomEditExample (原地編輯)...");

            var templateFile = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Template.xlsx";
            var editedFile = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Edited.xlsx";

            // 1. 準備一個範本檔案 (包含一些假資料)
            var wb = new FluentWorkbook(new XSSFWorkbook());
            wb.UseSheet("Report")
              .SetCellPosition(ExcelCol.A, 1).SetValue("Title: Monthly Report")
              .SetCellPosition(ExcelCol.A, 2).SetValue("Generated: [DATE]")
              .SetCellPosition(ExcelCol.B, 5).SetValue("Data Area");

            wb.SaveToFile(templateFile).Close();

            // 2. 載入並修改 (Load -> Edit -> Save)
            // 這裡使用 ReadExcelFile，它會將整份檔案載入記憶體 (DOM)
            // 所以原本的 "Data Area" 會被保留，我們只修改我們觸碰的儲存格
            var editor = new FluentWorkbook(new XSSFWorkbook());
            editor.ReadExcelFile(templateFile);

            editor.UseSheet("Report")
                  // 修改標題
                  .SetCellPosition(ExcelCol.A, 1).SetValue("Title: Final Report 2024")
                  // 填入日期
                  .SetCellPosition(ExcelCol.A, 2).SetValue($"Generated: {DateTime.Now:yyyy-MM-dd}")
                  // 新增一些數據
                  .SetCellPosition(ExcelCol.A, 10).SetValue("Approved by Manager");

            editor.SaveToFile(editedFile);
            editor.Close();

            Console.WriteLine($"  ✓ DOM 編輯完成: {editedFile}");
        }

        /// <summary>
        /// Example 12: Export to HTML
        /// </summary>
        static void CreateHtmlExportExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 HtmlExportExample...");

            var htmlPath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\ExportedRequest.html";

            // 1. 建立一個展示用 Sheet，包含合併、顏色與字體
            Console.WriteLine("  > 正在建立 'HtmlDemo' Sheet 以展示樣式支援...");

            // 定義樣式
            fluent.SetupCellStyle("MergedTitle", (w, s) =>
            {
                s.SetAlignment(HorizontalAlignment.Center);
                s.SetFontInfo(w, fontFamily: "Arial", fontHeight: 16, isBold: true);
                s.FillPattern = FillPattern.SolidForeground;
                s.SetCellFillForegroundColor(IndexedColors.LightCornflowerBlue);
                s.SetBorderAllStyle(BorderStyle.Thick); // Add Thick Border
            });

            fluent.SetupCellStyle("RedBg", (w, s) =>
            {
                s.FillPattern = FillPattern.SolidForeground;
                s.SetCellFillForegroundColor(IndexedColors.Red);
                s.SetFontInfo(w, color: IndexedColors.White);
                s.SetAlignment(HorizontalAlignment.Center);
            });

            fluent.SetupCellStyle("GreenText", (w, s) =>
            {
                s.SetFontInfo(w, color: IndexedColors.Green, isItalic: true, isBold: true);
                s.SetBorderAllStyle(BorderStyle.Dotted); // Add Dotted Border
            });

            fluent.SetupCellStyle("NumberFmt", (w, s) =>
            {
                s.SetDataFormat(w, "#,##0.00");
            });

            // 建立 Sheet 內容
            var sheet = fluent.UseSheet("HtmlDemo", true);

            // A1: 合併標題
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("HTML Export Feature Demo")
                 .SetCellStyle("MergedTitle");

            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1, 1);

            // A2: 紅色背景
            sheet.SetCellPosition(ExcelCol.A, 2).SetValue("Red Background")
                 .SetCellStyle("RedBg");

            // B2: 綠色文字
            sheet.SetCellPosition(ExcelCol.B, 2).SetValue("Green Italic Text")
                 .SetCellStyle("GreenText");

            // C2: 普通數值
            sheet.SetCellPosition(ExcelCol.C, 2).SetValue(1234.56)
                 .SetCellStyle("NumberFmt");

            // 2. 匯出為 HTML
            fluent.SaveAsHtml(htmlPath, fullHtml: true);

            // 3. 取得 HTML 字串 (僅表格片段)
            var htmlFragment = fluent.ToHtmlString(fullHtml: false);

            Console.WriteLine($"  ✓ HTML 匯出完成: {htmlPath}");
            Console.WriteLine($"  ✓ HTML 片段預覽 (前 100 字): {htmlFragment.Substring(0, Math.Min(100, htmlFragment.Length))}...");
        }

        /// <summary>
        /// Example 13: Export to PDF
        /// </summary>
        static void CreatePdfExportExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 PdfExportExample...");

            var pdfPath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\ExportedReport.pdf";

            // 使用剛建立的 HtmlDemo Sheet (已有樣式)
            fluent.UseSheet("HtmlDemo", true);
            fluent.SaveAsPdf(pdfPath);

            Console.WriteLine($"  ✓ PDF 匯出完成: {pdfPath}");
        }

        #endregion




        #region Read Examples

        static void ReadExcelExamples(FluentWorkbook fluent)
        {
            Console.WriteLine("\n========== 讀取 Excel 資料 ==========");

            // 讀取 BasicTableExample
            var sheet1 = fluent.UseSheet("BasicTableExample");
            Console.WriteLine("\n【BasicTableExample 標題行】:");
            for (ExcelCol col = ExcelCol.A; col <= ExcelCol.H; col++)
            {
                var headerValue = sheet1.GetCellValue<string>(col, 1);
                Console.Write($"{headerValue}\t");
            }
            Console.WriteLine();

            Console.WriteLine("\n【Sheet1 前3筆資料】:");
            for (int row = 2; row <= 4; row++)
            {
                var id = sheet1.GetCellValue<int>(ExcelCol.A, row);
                var name = sheet1.GetCellValue<string>(ExcelCol.B, row);
                var dateOfBirth = sheet1.GetCellValue<DateTime>(ExcelCol.C, row);
                Console.WriteLine($"Row {row}: ID={id}, Name={name}, Birth={dateOfBirth:yyyy-MM-dd}");
            }

            Console.WriteLine("\n========== 讀取完成 ==========\n");
        }

        #endregion
    }
}
