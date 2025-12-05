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

                // 設置樣式
                SetupStyles(fluent);

                // 創建各種範例工作表
                CreateBasicTableExample(fluent, testData);
                CreateSummaryExample(fluent, testData);
                CreateCopyStyleExample(fluent, testData);
                CreateDataTableExample(fluent);
                CreateCellStyleRangeExample(fluent);
                CreateSetCellValueExample(fluent);
                CreateCellMergeExample(fluent);
                CreatePictureExample(fluent);

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

        #region Write Examples

        /// <summary>
        /// 範例1：基本表格 - 展示多種欄位型別
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
        /// 範例2：摘要表格 - 展示不同欄位組合
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

        /// <summary>
        /// 範例3：複製樣式與動態樣式
        /// </summary>
        static void CreateCopyStyleExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            fluent.UseSheet("CopyStyleExample", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.D, 20)
                .SetTable(testData, ExcelCol.A, 1)
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
                .BeginTitleSet("姓名").CopyStyleFromCell(ExcelCol.A, 1)
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

        /// <summary>
        /// 範例6：設置單個單元格值並套用樣式
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
        /// 範例7：合併儲存格示例
        /// </summary>
        static void CreateCellMergeExample(FluentWorkbook fluent)
        {
            var sheet = fluent.UseSheet("CellMergeExample", true);

            // 設置欄寬
            sheet.SetColumnWidth(ExcelCol.A, ExcelCol.E, 15);

            // 1. 橫向合併（標題行）
            sheet.SetCellPosition(ExcelCol.A, 1)
                .SetValue("銷售報表")
                .SetCellStyle("HeaderBlue");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.E, 1); // 合併 A1-E1

            // 2. 設置子標題
            sheet.SetCellPosition(ExcelCol.A, 2).SetValue("產品名稱");
            sheet.SetCellPosition(ExcelCol.B, 2).SetValue("銷售量");
            sheet.SetCellPosition(ExcelCol.C, 2).SetValue("單價");
            sheet.SetCellPosition(ExcelCol.D, 2).SetValue("總金額");
            sheet.SetCellPosition(ExcelCol.E, 2).SetValue("備註");

            // 為標題行設置樣式
            for (ExcelCol col = ExcelCol.A; col <= ExcelCol.E; col++)
            {
                sheet.SetCellPosition(col, 2).SetCellStyle("HeaderBlue");
            }

            // 3. 縱向合併（用於分類）
            sheet.SetCellPosition(ExcelCol.A, 3).SetValue("電子產品");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 3, 5); // 合併 A3-A5

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

            // 4. 區域合併（用於總計）
            sheet.SetCellPosition(ExcelCol.A, 6).SetValue("總計");
            sheet.SetCellPosition(ExcelCol.D, 6).SetValue(840000);
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 6); // 合併 A6-C6（橫向）
            sheet.SetCellPosition(ExcelCol.A, 6).SetCellStyle("HighlightYellow");

            // 5. 多個合併區域示例
            sheet.SetCellPosition(ExcelCol.A, 8).SetValue("部門A");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 8, 10); // 縱向合併 A8-A10

            sheet.SetCellPosition(ExcelCol.B, 8).SetValue("部門B");
            sheet.SetExcelCellMerge(ExcelCol.B, ExcelCol.B, 8, 10); // 縱向合併 B8-B10

            sheet.SetCellPosition(ExcelCol.C, 8).SetValue("部門C");
            sheet.SetExcelCellMerge(ExcelCol.C, ExcelCol.C, 8, 10); // 縱向合併 C8-C10

            // 6. 區域合併示例（多行多列）
            sheet.SetCellPosition(ExcelCol.A, 12).SetValue("重要通知");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.E, 12, 14); // 合併 A12-E14（區域）
            sheet.SetCellPosition(ExcelCol.A, 12).SetCellStyle("HighlightYellow");
        }

        /// <summary>
        /// 範例7：插入圖片示例
        /// </summary>
        static void CreatePictureExample(FluentWorkbook fluent)
        {
            var sheet = fluent.UseSheet("PictureExample", true);

            // 設置欄寬
            sheet.SetColumnWidth(ExcelCol.A, ExcelCol.D, 20);

            // 讀取圖片文件
            var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "pain.jpg");
            
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"警告：圖片文件不存在: {imagePath}");
                return;
            }

            byte[] imageBytes = File.ReadAllBytes(imagePath);

            // 1. 基本插入圖片（自動計算高度，使用默認列寬比例）
            sheet.SetCellPosition(ExcelCol.A, 1)
                .SetValue("基本插入（自動高度）")
                .SetCellStyle("HeaderBlue");
            
            sheet.SetCellPosition(ExcelCol.A, 2)
                .SetPictureOnCell(imageBytes, 200); // 寬度 200 像素，高度自動計算（1:1）

            // 2. 手動設置寬度和高度
            sheet.SetCellPosition(ExcelCol.B, 1)
                .SetValue("手動設置尺寸")
                .SetCellStyle("HeaderBlue");
            
            sheet.SetCellPosition(ExcelCol.B, 2)
                .SetPictureOnCell(imageBytes, 200, 150); // 寬度 200，高度 150 像素

            // 3. 自定義列寬轉換比例
            sheet.SetCellPosition(ExcelCol.C, 1)
                .SetValue("自定義列寬比例")
                .SetCellStyle("HeaderBlue");
            
            sheet.SetCellPosition(ExcelCol.C, 2)
                .SetPictureOnCell(imageBytes, 300, AnchorType.MoveAndResize, 5.0); // 使用 5.0 作為轉換比例

            // 4. 鏈式調用示例（插入圖片後繼續設置樣式）
            sheet.SetCellPosition(ExcelCol.D, 1)
                .SetValue("鏈式調用示例")
                .SetCellStyle("HeaderBlue");
            
            sheet.SetCellPosition(ExcelCol.D, 2)
                .SetPictureOnCell(imageBytes, 180, 180, AnchorType.MoveAndResize, 7.0)
                .SetValue("圖片下方文字"); // 鏈式調用，在圖片後設置文字

            // 5. 不同錨點類型示例
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

            // 6. 多個圖片排列示例
            sheet.SetCellPosition(ExcelCol.A, 9)
                .SetValue("多圖片排列")
                .SetCellStyle("HeaderBlue");

            for (int i = 0; i < 3; i++)
            {
                sheet.SetCellPosition((ExcelCol)((int)ExcelCol.A + i), 10)
                    .SetPictureOnCell(imageBytes, 100, 100);
            }

            Console.WriteLine("✓ 圖片插入示例已創建");
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
            ReadGetTableExample(fluent);
            ReadCellMergeExample(fluent);

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
            for (ExcelCol col = ExcelCol.A; col <= ExcelCol.H; col++)
            {
                var headerValue = sheet1.GetCellValue<string>(col, 1);
                Console.Write($"{headerValue}\t");
            }
            Console.WriteLine();

            // 讀取資料行（第2行開始）
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
        /// 讀取範例2：讀取 DataTableExample 的資料
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
        /// 讀取範例3：使用 FluentCell 讀取單個單元格
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
        /// 讀取範例4：讀取 SetCellValueExample
        /// </summary>
        static void ReadSetCellValueExample(FluentWorkbook fluent)
        {
            Console.WriteLine("\n【SetCellValueExample 示例】:");
            var exampleSheet = fluent.UseSheet("SetCellValueExample");
            var helloValue = exampleSheet.GetCellValue<string>(ExcelCol.A, 1);
            Console.WriteLine($"A1 值: {helloValue}");
        }

        /// <summary>
        /// 讀取範例5：使用 GetTable<T> 讀取整個表格
        /// </summary>
        static void ReadGetTableExample(FluentWorkbook fluent)
        {
            Console.WriteLine("\n【使用 GetTable<T> 讀取表格】:");
            var sheet1 = fluent.UseSheet("Sheet1");

            // 方法1：指定結束行（舊方法，仍然可用）
            Console.WriteLine("\n方法1：指定結束行");
            var readData1 = sheet1.GetTable<ExampleData>(ExcelCol.A, 2, 13);
            Console.WriteLine($"成功讀取 {readData1.Count} 筆資料（指定結束行）");

            // 方法2：自動判斷最後一行（新方法）
            Console.WriteLine("\n方法2：自動判斷最後一行");
            var readData2 = sheet1.GetTable<ExampleData>(ExcelCol.A, 2);
            Console.WriteLine($"成功讀取 {readData2.Count} 筆資料（自動判斷最後一行）");

            // 方法3：指定列範圍和行範圍（新方法）
            Console.WriteLine("\n方法3：指定列範圍和行範圍");
            // 創建一個只包含部分欄位的類來演示
            var readData3 = sheet1.GetTable<PartialData>(ExcelCol.A, ExcelCol.C, 2, 5);
            Console.WriteLine($"成功讀取 {readData3.Count} 筆資料（只讀取 A-C 列，第 2-5 行）");
            Console.WriteLine("前3筆資料詳情（只包含 ID, Name, DateOfBirth）:");
            for (int i = 0; i < Math.Min(3, readData3.Count); i++)
            {
                var item = readData3[i];
                Console.WriteLine($"  [{i + 1}] ID={item.ID}, Name={item.Name}, Birth={item.DateOfBirth:yyyy-MM-dd}");
            }

            // 方法3示例2：讀取中間的列範圍
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

        /// <summary>
        /// 讀取範例6：讀取合併儲存格示例
        /// </summary>
        static void ReadCellMergeExample(FluentWorkbook fluent)
        {
            Console.WriteLine("\n【CellMergeExample 合併儲存格示例】:");
            var mergeSheet = fluent.UseSheet("CellMergeExample");
            var npoiSheet = mergeSheet.GetSheet();

            // 顯示合併區域資訊
            Console.WriteLine($"\n工作表共有 {npoiSheet.NumMergedRegions} 個合併區域：");

            for (int i = 0; i < npoiSheet.NumMergedRegions; i++)
            {
                var mergedRegion = npoiSheet.GetMergedRegion(i);
                var firstRow = mergedRegion.FirstRow + 1; // 轉換為 1-based
                var lastRow = mergedRegion.LastRow + 1;
                var firstCol = (ExcelCol)mergedRegion.FirstColumn;
                var lastCol = (ExcelCol)mergedRegion.LastColumn;

                // 讀取合併區域的第一個單元格的值
                var cellValue = mergeSheet.GetCellValue<string>(firstCol, firstRow) ?? "";

                if (firstRow == lastRow && firstCol == lastCol)
                {
                    // 單個單元格（不應該發生，但以防萬一）
                    Console.WriteLine($"  [{i + 1}] {firstCol}{firstRow}: {cellValue}");
                }
                else if (firstRow == lastRow)
                {
                    // 橫向合併
                    Console.WriteLine($"  [{i + 1}] 橫向合併: {firstCol}{firstRow}-{lastCol}{firstRow} = \"{cellValue}\"");
                }
                else if (firstCol == lastCol)
                {
                    // 縱向合併
                    Console.WriteLine($"  [{i + 1}] 縱向合併: {firstCol}{firstRow}-{firstCol}{lastRow} = \"{cellValue}\"");
                }
                else
                {
                    // 區域合併
                    Console.WriteLine($"  [{i + 1}] 區域合併: {firstCol}{firstRow}-{lastCol}{lastRow} = \"{cellValue}\"");
                }
            }

            // 讀取一些具體的合併單元格值
            Console.WriteLine("\n讀取合併單元格的值：");
            Console.WriteLine($"  A1 (合併區域): {mergeSheet.GetCellValue<string>(ExcelCol.A, 1)}");
            Console.WriteLine($"  A3 (縱向合併): {mergeSheet.GetCellValue<string>(ExcelCol.A, 3)}");
            Console.WriteLine($"  A6 (橫向合併): {mergeSheet.GetCellValue<string>(ExcelCol.A, 6)}");
            Console.WriteLine($"  A12 (區域合併): {mergeSheet.GetCellValue<string>(ExcelCol.A, 12)}");
        }

        #endregion
    }

    /// <summary>
    /// 部分欄位數據類（用於演示只讀取部分列）
    /// </summary>
    public class PartialData
    {
        public int ID { get; set; }
        public string Name { get; set; } = string.Empty;
        public DateTime DateOfBirth { get; set; }
    }

    /// <summary>
    /// 中間欄位數據類（用於演示讀取中間的列範圍）
    /// </summary>
    public class MiddleColumnsData
    {
        public bool IsActive { get; set; }
        public double Score { get; set; }
        public decimal Amount { get; set; }
    }

}
