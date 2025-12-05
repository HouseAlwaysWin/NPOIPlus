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
                var outputPath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Test2.xlsx";

                var fluent = new FluentWorkbook(new XSSFWorkbook(filePath));

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

                // Cell write examples
                CreateSetCellValueExample(fluent);
                CreateCellMergeExample(fluent);
                CreatePictureExample(fluent);

                // Save file
                fluent.SaveToPath(outputPath);

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
                .SetupGlobalCachedCellStyles((workbook, style) =>
                {
                    style.SetAligment(HorizontalAlignment.Center);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetFontInfo(workbook, "Calibri", 10);
                })
                .SetupCellStyle("DateOfBirth", (workbook, style) =>
                {
                    style.SetDataFormat(workbook, "yyyy-MM-dd");
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetAligment(HorizontalAlignment.Center);
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                })
                .SetupCellStyle("HeaderBlue", (workbook, style) =>
                {
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetAligment(HorizontalAlignment.Center);
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightCornflowerBlue);
                })
                .SetupCellStyle("BodyGreen", (workbook, style) =>
                {
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetAligment(HorizontalAlignment.Center);
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                })
                .SetupCellStyle("AmountCurrency", (workbook, style) =>
                {
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetDataFormat(workbook, "#,##0.00");
                    style.SetAligment(HorizontalAlignment.Right);
                })
                .SetupCellStyle("HighlightYellow", (workbook, style) =>
                {
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Yellow);
                });
        }

        #endregion

        #region Table Write Examples

        /// <summary>
        /// Example 1: Basic table - Demonstrates various field types
        /// </summary>
        static void CreateBasicTableExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            fluent.UseSheet("Sheet1")
                .SetColumnWidth(ExcelCol.A, ExcelCol.H, 20)
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID").SetCellStyle("HeaderBlue")
                .BeginBodySet("ID").SetCellStyle("BodyGreen").End()

                .BeginTitleSet("名稱").SetCellStyle("HeaderBlue")
                .BeginBodySet("Name").SetCellStyle("BodyGreen").End()

                .BeginTitleSet("生日").SetCellStyle("HeaderBlue")
                .BeginBodySet("DateOfBirth").SetCellStyle("DateOfBirth").End()

                .BeginTitleSet("是否活躍").SetCellStyle("HeaderBlue")
                .BeginBodySet("IsActive").SetCellType(CellType.Boolean).End()

                .BeginTitleSet("分數").SetCellStyle("HeaderBlue")
                .BeginBodySet("Score").SetCellType(CellType.Numeric).End()

                .BeginTitleSet("金額").SetCellStyle("HeaderBlue")
                .BeginBodySet("Amount").SetCellType(CellType.Numeric).SetCellStyle("AmountCurrency").End()

                .BeginTitleSet("備註").SetCellStyle("HeaderBlue")
                .BeginBodySet("Notes").SetCellType(CellType.String).End()

                .BeginTitleSet("可能為空").SetCellStyle("HeaderBlue")
                .BeginBodySet("MaybeNull").End()
                .BuildRows();
        }

        /// <summary>
        /// Example 2: Summary table - Demonstrates different field combinations
        /// </summary>
        static void CreateSummaryExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            fluent.UseSheet("Summary", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.E, 20)
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("姓名").SetCellStyle("HeaderBlue")
                .BeginBodySet("Name").End()

                .BeginTitleSet("分數").SetCellStyle("HeaderBlue")
                .BeginBodySet("Score").SetCellType(CellType.Numeric).SetCellStyle("AmountCurrency").End()

                .BeginTitleSet("日期").SetCellStyle("HeaderBlue")
                .BeginBodySet("DateOfBirth").SetCellStyle("DateOfBirth").End()

                .BeginTitleSet("狀態").SetCellStyle("HeaderBlue")
                .BeginBodySet("IsActive").SetCellType(CellType.Boolean).End()

                .BeginTitleSet("備註").SetCellStyle("HeaderBlue")
                .BeginBodySet("Notes").SetCellStyle("HighlightYellow").End()
                .BuildRows();
        }

        #endregion

        #region Style Write Examples

        /// <summary>
        /// Example 3: Copy style and dynamic style
        /// </summary>
        static void CreateCopyStyleExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            fluent.UseSheet("CopyStyleExample", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.D, 20)
                .SetTable(testData, ExcelCol.A, 1)
                // Demonstrate ID field with custom style
                .BeginTitleSet("編號").SetCellStyle("HeaderBlue")
                .BeginBodySet("ID")
                .SetCellStyle((styleParams) =>
                {
                    return new CellStyleConfig("IDStyle", style =>
                    {
                        style.SetAligment(HorizontalAlignment.Center);
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    });
                })
                .End()

                // Copy style from Sheet1
                .BeginTitleSet("姓名").CopyStyleFromCell(ExcelCol.A, 1)
                .BeginBodySet("Name").SetCellType(CellType.String).End()

                // Demonstrate amount field with currency format
                .BeginTitleSet("金額").SetCellStyle("HeaderBlue")
                .BeginBodySet("Amount").SetCellType(CellType.Numeric)
                .SetCellStyle("AmountCurrency")
                .End()

                // Demonstrate active status field (dynamic style based on data)
                .BeginTitleSet("活躍").SetCellStyle("HeaderBlue")
                .BeginBodySet("IsActive").SetCellType(CellType.Boolean)
                .SetCellStyle((styleParams) =>
                {
                    // Determine which style to return based on data
                    if (styleParams.GetRowItem<ExampleData>().IsActive)
                    {
                        return new("IsActiveStyle1", style =>
                        {
                            style.SetAligment(HorizontalAlignment.Center);
                            style.FillPattern = FillPattern.SolidForeground;
                            style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                            style.SetBorderAllStyle(BorderStyle.Thin);
                        });
                    }
                    return new("IsActiveStyle2", style =>
                    {
                        style.SetAligment(HorizontalAlignment.Center);
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor(IndexedColors.LightYellow);
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    });
                })
                .End()
                .BuildRows();
        }

        /// <summary>
        /// Example 4: Using DataTable as data source
        /// </summary>
        static void CreateDataTableExample(FluentWorkbook fluent)
        {
            DataTable dataTable = new DataTable("StudentData");
            dataTable.Columns.Add("StudentID", typeof(int));
            dataTable.Columns.Add("StudentName", typeof(string));
            dataTable.Columns.Add("BirthDate", typeof(DateTime));
            dataTable.Columns.Add("IsEnrolled", typeof(bool));
            dataTable.Columns.Add("GPA", typeof(double));
            dataTable.Columns.Add("Tuition", typeof(decimal));
            dataTable.Columns.Add("Department", typeof(string));

            // Add data
            dataTable.Rows.Add(101, "張三", new DateTime(1998, 3, 15), true, 3.8, 25000m, "資訊工程");
            dataTable.Rows.Add(102, "李四", new DateTime(1999, 7, 20), true, 3.5, 22000m, "電機工程");
            dataTable.Rows.Add(103, "王五", new DateTime(1997, 11, 5), false, 2.9, 20000m, "機械工程");
            dataTable.Rows.Add(104, "趙六", new DateTime(2000, 1, 30), true, 3.9, 28000m, "資訊工程");
            dataTable.Rows.Add(105, "陳七", new DateTime(1998, 9, 12), true, 3.6, 23000m, "企業管理");
            dataTable.Rows.Add(106, "林八", new DateTime(1999, 5, 8), false, 3.2, 21000m, "財務金融");

            fluent.UseSheet("DataTableExample", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.G, 20)
                .SetTable<DataRow>(dataTable.Rows.Cast<DataRow>(), ExcelCol.A, 1)

                .BeginTitleSet("學號").SetCellStyle("HeaderBlue")
                .BeginBodySet("StudentID").SetCellType(CellType.Numeric).End()

                .BeginTitleSet("姓名").SetCellStyle("HeaderBlue")
                .BeginBodySet("StudentName").SetCellType(CellType.String).End()

                .BeginTitleSet("出生日期").SetCellStyle("HeaderBlue")
                .BeginBodySet("BirthDate").SetCellStyle("DateOfBirth").End()

                .BeginTitleSet("在學中").SetCellStyle("HeaderBlue")
                .BeginBodySet("IsEnrolled").SetCellType(CellType.Boolean).End()

                .BeginTitleSet("GPA").SetCellStyle("HeaderBlue")
                .BeginBodySet("GPA").SetCellType(CellType.Numeric).End()

                .BeginTitleSet("學費").SetCellStyle("HeaderBlue")
                .BeginBodySet("Tuition").SetCellType(CellType.Numeric).SetCellStyle("AmountCurrency").End()

                .BeginTitleSet("科系").SetCellStyle("HeaderBlue")
                .BeginBodySet("Department").SetCellType(CellType.String)
                .SetCellStyle((styleParams) =>
                {
                    // Dynamically set color based on department
                    var row = styleParams.RowItem as DataRow;
                    var dept = row?["Department"]?.ToString() ?? "";

                    if (dept == "資訊工程")
                    {
                        return new("DeptIT", style =>
                        {
                            style.SetAligment(HorizontalAlignment.Center);
                            style.FillPattern = FillPattern.SolidForeground;
                            style.SetCellFillForegroundColor(IndexedColors.LightBlue);
                            style.SetBorderAllStyle(BorderStyle.Thin);
                        });
                    }
                    else if (dept == "電機工程")
                    {
                        return new("DeptEE", style =>
                        {
                            style.SetAligment(HorizontalAlignment.Center);
                            style.FillPattern = FillPattern.SolidForeground;
                            style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                            style.SetBorderAllStyle(BorderStyle.Thin);
                        });
                    }
                    return new("DeptOther", style =>
                    {
                        style.SetAligment(HorizontalAlignment.Center);
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor(IndexedColors.Grey25Percent);
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    });
                })
                .End()
                .BuildRows();
        }

        /// <summary>
        /// Example 5: Batch set cell range styles
        /// </summary>
        static void CreateCellStyleRangeExample(FluentWorkbook fluent)
        {
            fluent.UseSheet("CellStyleRangeDemo")
                .SetCellStyleRange(new CellStyleConfig("HighlightRed", style =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Red);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                }), ExcelCol.A, ExcelCol.D, 1, 3)
                .SetCellStyleRange(new CellStyleConfig("HighlightOrange",
                    style =>
                    {
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor(IndexedColors.Orange);
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    }), ExcelCol.A, ExcelCol.D, 4, 6)
                .SetCellStyleRange(new CellStyleConfig("HighlightYellow",
                    style =>
                    {
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor(IndexedColors.Yellow);
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    }), ExcelCol.A, ExcelCol.D, 7, 9)
                .SetCellStyleRange(new CellStyleConfig("HighlightGreen",
                    style =>
                    {
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor(IndexedColors.Green);
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    }), ExcelCol.A, ExcelCol.D, 10, 12)
                .SetCellStyleRange(new CellStyleConfig("HighlightBlue",
                    style =>
                    {
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor(IndexedColors.Blue);
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    }), ExcelCol.A, ExcelCol.D, 13, 15)
                .SetCellStyleRange(new CellStyleConfig("HighlightPurple",
                    style =>
                    {
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor("#FF00FF");
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    }), ExcelCol.A, ExcelCol.D, 16, 18);
        }

        #endregion

        #region Cell Write Examples

        /// <summary>
        /// Example 6: Set single cell value and apply style
        /// </summary>
        static void CreateSetCellValueExample(FluentWorkbook fluent)
        {
            fluent.UseSheet("SetCellValueExample", true)
                .SetColumnWidth(ExcelCol.A, 20)
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Hello, World!")
                .SetCellStyle("HighlightYellow");
        }

        /// <summary>
        /// Example 7: Cell merge example
        /// </summary>
        static void CreateCellMergeExample(FluentWorkbook fluent)
        {
            var sheet = fluent.UseSheet("CellMergeExample", true);

            // Set column width
            sheet.SetColumnWidth(ExcelCol.A, ExcelCol.E, 15);

            // 1. Horizontal merge (header row)
            sheet.SetCellPosition(ExcelCol.A, 1)
                .SetValue("銷售報表")
                .SetCellStyle("HeaderBlue");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.E, 1); // Merge A1-E1

            // 2. Set sub-headers
            sheet.SetCellPosition(ExcelCol.A, 2).SetValue("產品名稱");
            sheet.SetCellPosition(ExcelCol.B, 2).SetValue("銷售量");
            sheet.SetCellPosition(ExcelCol.C, 2).SetValue("單價");
            sheet.SetCellPosition(ExcelCol.D, 2).SetValue("總金額");
            sheet.SetCellPosition(ExcelCol.E, 2).SetValue("備註");

            // Apply style to header row
            for (ExcelCol col = ExcelCol.A; col <= ExcelCol.E; col++)
            {
                sheet.SetCellPosition(col, 2).SetCellStyle("HeaderBlue");
            }

            // 3. Vertical merge (for categorization)
            sheet.SetCellPosition(ExcelCol.A, 3).SetValue("電子產品");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 3, 5); // Merge A3-A5

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

            // 4. Region merge (for totals)
            sheet.SetCellPosition(ExcelCol.A, 6).SetValue("總計");
            sheet.SetCellPosition(ExcelCol.D, 6).SetValue(840000);
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 6); // Merge A6-C6 (horizontal)
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
        }

        /// <summary>
        /// Example 8: Insert picture example
        /// </summary>
        static void CreatePictureExample(FluentWorkbook fluent)
        {
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

            Console.WriteLine("✓ Picture insertion example created");
        }

        #endregion

        #region Table Read Examples

        static void ReadExcelExamples(FluentWorkbook fluent)
        {
            Console.WriteLine("\n========== Read Excel Data Examples ==========");

            // Table read examples
            ReadSheet1Example(fluent);
            ReadDataTableExample(fluent);
            ReadGetTableExample(fluent);

            // Cell read examples
            ReadFluentCellExample(fluent);
            ReadSetCellValueExample(fluent);
            ReadCellMergeExample(fluent);

            Console.WriteLine("\n========== Read Complete ==========\n");
        }

        /// <summary>
        /// Read Example 1: Read data from Sheet1
        /// </summary>
        static void ReadSheet1Example(FluentWorkbook fluent)
        {
            var sheet1 = fluent.UseSheet("Sheet1");

            // Read header row
            Console.WriteLine("\n【Sheet1 標題行】:");
            for (ExcelCol col = ExcelCol.A; col <= ExcelCol.H; col++)
            {
                var headerValue = sheet1.GetCellValue<string>(col, 1);
                Console.Write($"{headerValue}\t");
            }
            Console.WriteLine();

            // Read data rows (starting from row 2)
            Console.WriteLine("\n【Sheet1 前3筆資料】:");
            for (int row = 2; row <= 4; row++)
            {
                var id = sheet1.GetCellValue<int>(ExcelCol.A, row);
                var name = sheet1.GetCellValue<string>(ExcelCol.B, row);
                var dateOfBirth = sheet1.GetCellValue<DateTime>(ExcelCol.C, row);
                var isActive = sheet1.GetCellValue<bool>(ExcelCol.D, row);
                var score = sheet1.GetCellValue<double>(ExcelCol.E, row);
                var amount = sheet1.GetCellValue<double>(ExcelCol.F, row);
                var notes = sheet1.GetCellValue<string>(ExcelCol.G, row);

                Console.WriteLine($"Row {row}: ID={id}, Name={name}, Birth={dateOfBirth:yyyy-MM-dd}, Active={isActive}, Score={score}, Amount={amount:C}, Notes={notes}");
            }
        }

        /// <summary>
        /// Read Example 2: Read data from DataTableExample
        /// </summary>
        static void ReadDataTableExample(FluentWorkbook fluent)
        {
            Console.WriteLine("\n【DataTableExample 前3筆資料】:");
            var dtSheet = fluent.UseSheet("DataTableExample");
            for (int row = 2; row <= 4; row++)
            {
                var studentId = dtSheet.GetCellValue<int>(ExcelCol.A, row);
                var studentName = dtSheet.GetCellValue<string>(ExcelCol.B, row);
                var birthDate = dtSheet.GetCellValue<DateTime>(ExcelCol.C, row);
                var isEnrolled = dtSheet.GetCellValue<bool>(ExcelCol.D, row);
                var gpa = dtSheet.GetCellValue<double>(ExcelCol.E, row);
                var tuition = dtSheet.GetCellValue<double>(ExcelCol.F, row);
                var department = dtSheet.GetCellValue<string>(ExcelCol.G, row);

                Console.WriteLine($"Student {studentId}: {studentName}, {department}, GPA={gpa:F1}, 學費={tuition:C}");
            }
        }

        /// <summary>
        /// Read Example 3: Use GetTable<T> to read entire table
        /// </summary>
        static void ReadGetTableExample(FluentWorkbook fluent)
        {
            Console.WriteLine("\n【使用 GetTable<T> 讀取表格】:");
            var sheet1 = fluent.UseSheet("Sheet1");

            // Method 1: Specify end row (old method, still available)
            Console.WriteLine("\n方法1：指定結束行");
            var readData1 = sheet1.GetTable<ExampleData>(ExcelCol.A, 2, 13);
            Console.WriteLine($"成功讀取 {readData1.Count} 筆資料（指定結束行）");

            // Method 2: Auto-detect last row (new method)
            Console.WriteLine("\n方法2：自動判斷最後一行");
            var readData2 = sheet1.GetTable<ExampleData>(ExcelCol.A, 2);
            Console.WriteLine($"成功讀取 {readData2.Count} 筆資料（自動判斷最後一行）");

            // Method 3: Specify column and row range (new method)
            Console.WriteLine("\n方法3：指定列範圍和行範圍");
            // Create a class with only partial fields for demonstration
            var readData3 = sheet1.GetTable<PartialData>(ExcelCol.A, ExcelCol.C, 2, 5);
            Console.WriteLine($"成功讀取 {readData3.Count} 筆資料（只讀取 A-C 列，第 2-5 行）");
            Console.WriteLine("前3筆資料詳情（只包含 ID, Name, DateOfBirth）:");
            for (int i = 0; i < Math.Min(3, readData3.Count); i++)
            {
                var item = readData3[i];
                Console.WriteLine($"  [{i + 1}] ID={item.ID}, Name={item.Name}, Birth={item.DateOfBirth:yyyy-MM-dd}");
            }

            // Method 3 Example 2: Read middle column range
            Console.WriteLine("\n方法3示例2：讀取中間的列範圍（D-F 列，第 2-4 行）");
            var readData4 = sheet1.GetTable<MiddleColumnsData>(ExcelCol.D, ExcelCol.F, 2, 4);
            Console.WriteLine($"成功讀取 {readData4.Count} 筆資料");
            for (int i = 0; i < readData4.Count; i++)
            {
                var item = readData4[i];
                Console.WriteLine($"  [{i + 1}] IsActive={item.IsActive}, Score={item.Score:F1}, Amount={item.Amount:C}");
            }

            Console.WriteLine("\n前5筆完整資料詳情（方法2的結果）:");

            for (int i = 0; i < Math.Min(5, readData2.Count); i++)
            {
                var item = readData2[i];
                Console.WriteLine($"  [{i + 1}] ID={item.ID}, Name={item.Name}, " +
                    $"Birth={item.DateOfBirth:yyyy-MM-dd}, Active={item.IsActive}, " +
                    $"Score={item.Score:F1}, Amount={item.Amount:C}");
            }
        }

        #endregion

        #region Cell Read Examples

        /// <summary>
        /// Read Example 1: Use FluentCell to read single cell
        /// </summary>
        static void ReadFluentCellExample(FluentWorkbook fluent)
        {
            Console.WriteLine("\n【使用 FluentCell 讀取】:");
            var sheet1 = fluent.UseSheet("Sheet1");
            var cellA1 = sheet1.SetCellPosition(ExcelCol.A, 1);
            if (cellA1 != null)
            {
                var value = cellA1.GetValue();
                var cellType = cellA1.GetCell().CellType;
                Console.WriteLine($"A1 單元格: 值={value}, 類型={cellType}");
            }
        }

        /// <summary>
        /// Read Example 2: Read SetCellValueExample
        /// </summary>
        static void ReadSetCellValueExample(FluentWorkbook fluent)
        {
            Console.WriteLine("\n【SetCellValueExample 示例】:");
            var exampleSheet = fluent.UseSheet("SetCellValueExample");
            var helloValue = exampleSheet.GetCellValue<string>(ExcelCol.A, 1);
            Console.WriteLine($"A1 值: {helloValue}");
        }

        /// <summary>
        /// Read Example 3: Read merged cells example
        /// </summary>
        static void ReadCellMergeExample(FluentWorkbook fluent)
        {
            Console.WriteLine("\n【CellMergeExample 合併儲存格示例】:");
            var mergeSheet = fluent.UseSheet("CellMergeExample");
            var npoiSheet = mergeSheet.GetSheet();

            // Display merged region information
            Console.WriteLine($"\n工作表共有 {npoiSheet.NumMergedRegions} 個合併區域：");

            for (int i = 0; i < npoiSheet.NumMergedRegions; i++)
            {
                var mergedRegion = npoiSheet.GetMergedRegion(i);
                var firstRow = mergedRegion.FirstRow + 1; // Convert to 1-based
                var lastRow = mergedRegion.LastRow + 1;
                var firstCol = (ExcelCol)mergedRegion.FirstColumn;
                var lastCol = (ExcelCol)mergedRegion.LastColumn;

                // Read the value of the first cell in the merged region
                var cellValue = mergeSheet.GetCellValue<string>(firstCol, firstRow) ?? "";

                if (firstRow == lastRow && firstCol == lastCol)
                {
                    // Single cell (should not happen, but just in case)
                    Console.WriteLine($"  [{i + 1}] {firstCol}{firstRow}: {cellValue}");
                }
                else if (firstRow == lastRow)
                {
                    // Horizontal merge
                    Console.WriteLine($"  [{i + 1}] 橫向合併: {firstCol}{firstRow}-{lastCol}{firstRow} = \"{cellValue}\"");
                }
                else if (firstCol == lastCol)
                {
                    // Vertical merge
                    Console.WriteLine($"  [{i + 1}] 縱向合併: {firstCol}{firstRow}-{firstCol}{lastRow} = \"{cellValue}\"");
                }
                else
                {
                    // Region merge
                    Console.WriteLine($"  [{i + 1}] 區域合併: {firstCol}{firstRow}-{lastCol}{lastRow} = \"{cellValue}\"");
                }
            }

            // Read some specific merged cell values
            Console.WriteLine("\n讀取合併單元格的值：");
            Console.WriteLine($"  A1 (合併區域): {mergeSheet.GetCellValue<string>(ExcelCol.A, 1)}");
            Console.WriteLine($"  A3 (縱向合併): {mergeSheet.GetCellValue<string>(ExcelCol.A, 3)}");
            Console.WriteLine($"  A6 (橫向合併): {mergeSheet.GetCellValue<string>(ExcelCol.A, 6)}");
            Console.WriteLine($"  A12 (區域合併): {mergeSheet.GetCellValue<string>(ExcelCol.A, 12)}");
        }

        #endregion
    }

    /// <summary>
    /// Partial field data class (for demonstrating reading only partial columns)
    /// </summary>
    public class PartialData
    {
        public int ID { get; set; }
        public string Name { get; set; } = string.Empty;
        public DateTime DateOfBirth { get; set; }
    }

    /// <summary>
    /// Middle columns data class (for demonstrating reading middle column range)
    /// </summary>
    public class MiddleColumnsData
    {
        public bool IsActive { get; set; }
        public double Score { get; set; }
        public decimal Amount { get; set; }
    }

}
