using System;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Stages;
using FluentNPOI.HotReload;
using FluentNPOI.HotReload.Widgets;
using FluentNPOI.HotReload.Context;
using FluentNPOI.HotReload.Styling;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FluentNPOIConsoleExample;

/// <summary>
/// Hot Reload 範例程式
/// 使用方式: dotnet watch run
/// </summary>
public static class HotReloadExamples
{
    /// <summary>
    /// 範例 1: 使用 Widget 系統的宣告式風格
    /// </summary>
    public static void RunWidgetExample()
    {
        Console.WriteLine("=== Widget 系統範例 ===");
        Console.WriteLine("請用 dotnet watch run 執行以體驗熱重載功能");
        Console.WriteLine();

        ExcelLivePreview.Run<SalesReportWidget>("output/widget_demo.xlsx");
    }

    /// <summary>
    /// 範例 2: 使用現有 FluentNPOI 程式碼 (零學習成本)
    /// </summary>
    public static void RunFluentExample()
    {
        Console.WriteLine("=== FluentNPOI 原生風格範例 ===");
        Console.WriteLine("直接使用現有程式碼，無需學習 Widget 系統");
        Console.WriteLine();

        FluentLivePreview.Run("output/fluent_demo.xlsx", BuildFluentReport);
    }

    /// <summary>
    /// 範例 3: 混合使用 Widget 和 FluentNPOI
    /// </summary>
    public static void RunHybridExample()
    {
        Console.WriteLine("=== 混合模式範例 ===");
        Console.WriteLine("同時使用 Widget 和原生 FluentNPOI API");
        Console.WriteLine();

        FluentLivePreview.Run("output/hybrid_demo.xlsx", BuildHybridReport);
    }

    /// <summary>
    /// 現有 FluentNPOI 風格的報表
    /// 修改這裡的程式碼，儲存後會自動重新產生 Excel
    /// </summary>
    private static void BuildFluentReport(FluentWorkbook workbook)
    {
        var sheet = workbook.UseSheet("銷售報表");

        // 標題
        sheet.SetCellPosition(ExcelCol.A, 1)
            .SetValue("2024 年度銷售報表")
            .SetFont(isBold: true, fontSize: 16);

        // 欄位標題
        sheet.SetCellPosition(ExcelCol.A, 3).SetValue("產品").SetFont(isBold: true);
        sheet.SetCellPosition(ExcelCol.B, 3).SetValue("單價").SetFont(isBold: true);
        sheet.SetCellPosition(ExcelCol.C, 3).SetValue("數量").SetFont(isBold: true);
        sheet.SetCellPosition(ExcelCol.D, 3).SetValue("小計").SetFont(isBold: true);

        // 資料列 - 修改這裡試試熱重載！
        var products = new[]
        {
            ("蘋果", 30, 100),
            ("香蕉", 20, 150),
            ("橘子", 25, 80),
            ("葡萄", 50, 60),
        };

        for (int i = 0; i < products.Length; i++)
        {
            var (name, price, qty) = products[i];
            var row = 4 + i;

            sheet.SetCellPosition(ExcelCol.A, row).SetValue(name);
            sheet.SetCellPosition(ExcelCol.B, row).SetValue(price);
            sheet.SetCellPosition(ExcelCol.C, row).SetValue(qty);
            sheet.SetCellPosition(ExcelCol.D, row).SetValue(price * qty);
        }

        // 設定欄寬
        sheet.SetColumnWidth(ExcelCol.A, 15);
        sheet.SetColumnWidth(ExcelCol.B, 10);
        sheet.SetColumnWidth(ExcelCol.C, 10);
        sheet.SetColumnWidth(ExcelCol.D, 12);
    }

    /// <summary>
    /// 混合使用 Widget 和 FluentNPOI 的報表
    /// </summary>
    private static void BuildHybridReport(FluentWorkbook workbook)
    {
        var sheet = workbook.UseSheet("混合報表");

        // 使用 FluentNPOI 設定標題
        sheet.SetCellPosition(ExcelCol.A, 1)
            .SetValue("混合模式報表")
            .SetFont(isBold: true, fontSize: 14)
            .SetBackgroundColor(IndexedColors.LightBlue);

        // 使用 Widget 系統建立表格
        var tableWidget = new Column(
            new FlexibleRow(
                new Header("欄位 A").WithWeight(2),
                new Header("欄位 B").WithWeight(1),
                new Header("欄位 C").WithWeight(1)
            ).SetTotalWidth(60),
            new FlexibleRow(
                new Label("資料 1").WithWeight(2),
                new InputCell(100).WithWeight(1),
                new InputCell(200).WithWeight(1)
            ).SetTotalWidth(60),
            new FlexibleRow(
                new Label("資料 2").WithWeight(2),
                new InputCell(150).WithWeight(1),
                new InputCell(250).WithWeight(1)
            ).SetTotalWidth(60)
        );

        // 從第 3 行開始建立 Widget 表格
        sheet.BuildWidget(tableWidget, startRow: 3);

        // 繼續使用 FluentNPOI 加入註腳
        sheet.SetCellPosition(ExcelCol.A, 7)
            .SetValue("* 此報表使用混合模式產生")
            .SetFont(isBold: false);
    }
}

/// <summary>
/// Widget 風格的銷售報表
/// 修改這個類別的內容，儲存後會自動重新產生 Excel
/// </summary>
public class SalesReportWidget : ExcelWidget
{
    public override void Build(ExcelContext ctx)
    {
        // 建立完整的報表結構
        var report = new Column(
            // 標題區
            new Header("2024 年度銷售報表")
                .SetBackground(IndexedColors.RoyalBlue),

            // 空行
            new Label(""),

            // 表格標題
            new FlexibleRow(
                new Header("產品").WithWeight(2),
                new Header("單價").WithWeight(1),
                new Header("數量").WithWeight(1),
                new Header("小計").WithWeight(1)
            ).SetTotalWidth(60),

            // 資料列 - 修改這裡試試熱重載！
            CreateDataRow("蘋果", 30, 100),
            CreateDataRow("香蕉", 20, 150),
            CreateDataRow("橘子", 25, 80),
            CreateDataRow("葡萄", 50, 60),

            // 空行
            new Label(""),

            // 小計
            new FlexibleRow(
                new Label("總計:").WithWeight(2).WithStyle(new FluentStyle().SetBold()),
                new Label("").WithWeight(1),
                new Label("").WithWeight(1),
                new Label("=SUM(D4:D7)").WithWeight(1).WithStyle(new FluentStyle().SetBold())
            ).SetTotalWidth(60)
        );

        report.Build(ctx);
    }

    private static FlexibleRow CreateDataRow(string product, int price, int qty)
    {
        return new FlexibleRow(
            new Label(product).WithWeight(2),
            new Label(price.ToString()).WithWeight(1),
            new InputCell(qty).WithWeight(1),
            new Label((price * qty).ToString()).WithWeight(1)
        ).SetTotalWidth(60);
    }
}
