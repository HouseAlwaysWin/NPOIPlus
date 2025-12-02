using System.Data;
using System;
using System.IO;
using FluentNPOI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using FluentNPOI.Models;
using System.Linq;

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

                // 設置樣式
                SetupStyles(fluent);

                // 創建各種範例工作表
                CreateBasicTableExample(fluent, testData);
                CreateSummaryExample(fluent, testData);
                CreateCopyStyleExample(fluent, testData);
                CreateDataTableExample(fluent);
                CreateCellStyleRangeExample(fluent);
                CreateSetCellValueExample(fluent);

                // 儲存檔案
                fluent.SaveToPath(outputPath);

                // 讀取範例
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
                new ExampleData(1, "Alice Chen", new DateTime(1994, 1, 1), true, 95.5, 12500.75m, "優秀學生"),
                new ExampleData(2, "Bob Lee", new DateTime(1989, 5, 12), false, 78.0, 8900.50m, "需改進"),
                new ExampleData(3, "Søren", new DateTime(1985, 7, 23), true, 88.5, 15000.00m, "表現良好"),
                new ExampleData(4, "王小明", new DateTime(2000, 2, 29), true, 92.0, 11200.80m, "進步快速"),
                new ExampleData(5, "This is a very very very long name to test wrapping and width", new DateTime(1980, 12, 31), false, 65.5, 7500.25m, "LongName"),
                new ExampleData(6, "Élodie", new DateTime(1995, 5, 15), true, 85.0, 9800.00m, "穩定發揮"),
                new ExampleData(7, "O'Connor", new DateTime(1975, 7, 7), false, 72.5, 8200.50m, "待觀察"),
                new ExampleData(8, "李雷", new DateTime(2010, 10, 10), true, 90.0, 10500.75m, "潛力股"),
                new ExampleData(9, "山田太郎", new DateTime(1999, 3, 3), true, 87.5, 9500.00m, "穩健型"),
                new ExampleData(10, "Мария", new DateTime(1988, 8, 8), false, 70.0, 8000.25m, "需加強"),
                new ExampleData(11, "محمد", new DateTime(1991, 9, 9), true, 93.5, 12000.00m, "頂尖"),
                new ExampleData(12, "김민준", new DateTime(2004, 4, 4), true, 89.0, 10200.50m, "均衡發展"),
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

        #region Write Examples

        /// <summary>
        /// 範例1：基本表格 - 展示多種欄位型別
        /// </summary>
        static void CreateBasicTableExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            fluent.UseSheet("Sheet1")
                .SetColumnWidth(ExcelColumns.A, ExcelColumns.H, 20)
                .SetTable(testData, ExcelColumns.A, 1)
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
        /// 範例2：摘要表格 - 展示不同欄位組合
        /// </summary>
        static void CreateSummaryExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            fluent.UseSheet("Summary", true)
                .SetColumnWidth(ExcelColumns.A, ExcelColumns.E, 20)
                .SetTable(testData, ExcelColumns.A, 1)
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

        /// <summary>
        /// 範例3：複製樣式與動態樣式
        /// </summary>
        static void CreateCopyStyleExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            fluent.UseSheet("CopyStyleExample", true)
                .SetColumnWidth(ExcelColumns.A, ExcelColumns.D, 20)
                .SetTable(testData, ExcelColumns.A, 1)
                // 展示 ID 欄位並套用自訂樣式
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

                // 從 Sheet1 的樣式複製
                .BeginTitleSet("姓名").CopyStyleFromCell(ExcelColumns.A, 1)
                .BeginBodySet("Name").SetCellType(CellType.String).End()

                // 展示金額欄位並套用金額格式
                .BeginTitleSet("金額").SetCellStyle("HeaderBlue")
                .BeginBodySet("Amount").SetCellType(CellType.Numeric)
                .SetCellStyle("AmountCurrency")
                .End()

                // 展示活躍狀態欄位（根據數據動態變化樣式）
                .BeginTitleSet("活躍").SetCellStyle("HeaderBlue")
                .BeginBodySet("IsActive").SetCellType(CellType.Boolean)
                .SetCellStyle((styleParams) =>
                {
                    // ✅ 根據資料決定返回哪個樣式
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
        /// 範例4：使用 DataTable 作為資料來源
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

            // 添加資料
            dataTable.Rows.Add(101, "張三", new DateTime(1998, 3, 15), true, 3.8, 25000m, "資訊工程");
            dataTable.Rows.Add(102, "李四", new DateTime(1999, 7, 20), true, 3.5, 22000m, "電機工程");
            dataTable.Rows.Add(103, "王五", new DateTime(1997, 11, 5), false, 2.9, 20000m, "機械工程");
            dataTable.Rows.Add(104, "趙六", new DateTime(2000, 1, 30), true, 3.9, 28000m, "資訊工程");
            dataTable.Rows.Add(105, "陳七", new DateTime(1998, 9, 12), true, 3.6, 23000m, "企業管理");
            dataTable.Rows.Add(106, "林八", new DateTime(1999, 5, 8), false, 3.2, 21000m, "財務金融");

            fluent.UseSheet("DataTableExample", true)
                .SetColumnWidth(ExcelColumns.A, ExcelColumns.G, 20)
                .SetTable<DataRow>(dataTable.Rows.Cast<DataRow>(), ExcelColumns.A, 1)

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
                    // 根據科系動態設定顏色
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
        /// 範例5：批量設置單元格範圍樣式
        /// </summary>
        static void CreateCellStyleRangeExample(FluentWorkbook fluent)
        {
            fluent.UseSheet("CellStyleRangeDemo")
                .SetCellStyleRange(new CellStyleConfig("HighlightRed", style =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Red);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                }), ExcelColumns.A, ExcelColumns.D, 1, 3)
                .SetCellStyleRange(new CellStyleConfig("HighlightOrange",
                    style =>
                    {
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor(IndexedColors.Orange);
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    }), ExcelColumns.A, ExcelColumns.D, 4, 6)
                .SetCellStyleRange(new CellStyleConfig("HighlightYellow",
                    style =>
                    {
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor(IndexedColors.Yellow);
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    }), ExcelColumns.A, ExcelColumns.D, 7, 9)
                .SetCellStyleRange(new CellStyleConfig("HighlightGreen",
                    style =>
                    {
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor(IndexedColors.Green);
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    }), ExcelColumns.A, ExcelColumns.D, 10, 12)
                .SetCellStyleRange(new CellStyleConfig("HighlightBlue",
                    style =>
                    {
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor(IndexedColors.Blue);
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    }), ExcelColumns.A, ExcelColumns.D, 13, 15)
                .SetCellStyleRange(new CellStyleConfig("HighlightPurple",
                    style =>
                    {
                        style.FillPattern = FillPattern.SolidForeground;
                        style.SetCellFillForegroundColor("#FF00FF");
                        style.SetBorderAllStyle(BorderStyle.Thin);
                    }), ExcelColumns.A, ExcelColumns.D, 16, 18);


        }

        /// <summary>
        /// 範例6：設置單個單元格值並套用樣式
        /// </summary>
        static void CreateSetCellValueExample(FluentWorkbook fluent)
        {
            fluent.UseSheet("SetCellValueExample", true)
                .SetColumnWidth(ExcelColumns.A, 20)
                .SetCellPosition(ExcelColumns.A, 1)
                .SetValue("Hello, World!")
                .SetCellStyle("HighlightYellow");
        }

        #endregion

        #region Read Examples

        static void ReadExcelExamples(FluentWorkbook fluent)
        {
            Console.WriteLine("\n========== 讀取 Excel 數據示例 ==========");

            ReadSheet1Example(fluent);
            ReadDataTableExample(fluent);
            ReadFluentCellExample(fluent);
            ReadSetCellValueExample(fluent);

            Console.WriteLine("\n========== 讀取完成 ==========\n");
        }

        /// <summary>
        /// 讀取範例1：讀取 Sheet1 的資料
        /// </summary>
        static void ReadSheet1Example(FluentWorkbook fluent)
        {
            var sheet1 = fluent.UseSheet("Sheet1");

            // 讀取標題行
            Console.WriteLine("\n【Sheet1 標題行】:");
            for (ExcelColumns col = ExcelColumns.A; col <= ExcelColumns.H; col++)
            {
                var headerValue = sheet1.GetCellValue<string>(col, 1);
                Console.Write($"{headerValue}\t");
            }
            Console.WriteLine();

            // 讀取資料行（第2行開始）
            Console.WriteLine("\n【Sheet1 前3筆資料】:");
            for (int row = 2; row <= 4; row++)
            {
                var id = sheet1.GetCellValue<int>(ExcelColumns.A, row);
                var name = sheet1.GetCellValue<string>(ExcelColumns.B, row);
                var dateOfBirth = sheet1.GetCellValue<DateTime>(ExcelColumns.C, row);
                var isActive = sheet1.GetCellValue<bool>(ExcelColumns.D, row);
                var score = sheet1.GetCellValue<double>(ExcelColumns.E, row);
                var amount = sheet1.GetCellValue<double>(ExcelColumns.F, row);
                var notes = sheet1.GetCellValue<string>(ExcelColumns.G, row);

                Console.WriteLine($"Row {row}: ID={id}, Name={name}, Birth={dateOfBirth:yyyy-MM-dd}, Active={isActive}, Score={score}, Amount={amount:C}, Notes={notes}");
            }
        }

        /// <summary>
        /// 讀取範例2：讀取 DataTableExample 的資料
        /// </summary>
        static void ReadDataTableExample(FluentWorkbook fluent)
        {
            Console.WriteLine("\n【DataTableExample 前3筆資料】:");
            var dtSheet = fluent.UseSheet("DataTableExample");
            for (int row = 2; row <= 4; row++)
            {
                var studentId = dtSheet.GetCellValue<int>(ExcelColumns.A, row);
                var studentName = dtSheet.GetCellValue<string>(ExcelColumns.B, row);
                var birthDate = dtSheet.GetCellValue<DateTime>(ExcelColumns.C, row);
                var isEnrolled = dtSheet.GetCellValue<bool>(ExcelColumns.D, row);
                var gpa = dtSheet.GetCellValue<double>(ExcelColumns.E, row);
                var tuition = dtSheet.GetCellValue<double>(ExcelColumns.F, row);
                var department = dtSheet.GetCellValue<string>(ExcelColumns.G, row);

                Console.WriteLine($"Student {studentId}: {studentName}, {department}, GPA={gpa:F1}, 學費={tuition:C}");
            }
        }

        /// <summary>
        /// 讀取範例3：使用 FluentCell 讀取單個單元格
        /// </summary>
        static void ReadFluentCellExample(FluentWorkbook fluent)
        {
            Console.WriteLine("\n【使用 FluentCell 讀取】:");
            var sheet1 = fluent.UseSheet("Sheet1");
            var cellA1 = sheet1.GetCellPosition(ExcelColumns.A, 1);
            if (cellA1 != null)
            {
                var value = cellA1.GetValue();
                var cellType = cellA1.GetCell().CellType;
                Console.WriteLine($"A1 單元格: 值={value}, 類型={cellType}");
            }
        }

        /// <summary>
        /// 讀取範例4：讀取 SetCellValueExample
        /// </summary>
        static void ReadSetCellValueExample(FluentWorkbook fluent)
        {
            Console.WriteLine("\n【SetCellValueExample 示例】:");
            var exampleSheet = fluent.UseSheet("SetCellValueExample");
            var helloValue = exampleSheet.GetCellValue<string>(ExcelColumns.A, 1);
            Console.WriteLine($"A1 值: {helloValue}");
        }

        #endregion
    }

}
