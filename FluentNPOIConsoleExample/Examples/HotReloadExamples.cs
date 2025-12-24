using System;
using FluentNPOI.HotReload;

namespace FluentNPOIConsoleExample;

/// <summary>
/// Hot Reload 範例程式
/// 使用方式: dotnet watch run
/// </summary>
public static class HotReloadExamples
{


    /// <summary>
    /// Hot Reload Example: Wraps existing FluentNPOI code to provide live preview.
    /// Run with: dotnet watch run
    /// </summary>
    public static void RunHotReloadExample()
    {
        Console.WriteLine("=== Hot Reload Example ===");
        Console.WriteLine("Modify code in any Example file (e.g., TableExamples.cs) to see live changes.");
        Console.WriteLine();

        FluentLivePreview.Run("output/style_demo.xlsx", wb =>
        {
            // 先設定必要的樣式
            Program.SetupStyles(wb);

            // === 基礎範例 ===
            Program.CreateBasicTableExample(wb, Program.testData);
            Program.CreateSummaryExample(wb, Program.testData);
            Program.CreateDataTableExample(wb);

            // === 樣式範例 ===
            // Program.CreateCellStyleRangeExample(wb);
            // Program.CreateCopyStyleExample(wb, Program.testData); // (Part 2)
            // Program.CreateSheetGlobalStyleExample(wb, Program.testData); // (Part 3)
            // Program.CreateMappingStylingExample(wb, Program.testData); // (New Feature)

            // === 儲存格操作 ===
            // Program.CreateSetCellValueExample(wb);
            // Program.CreateCellMergeExample(wb);
            // Program.CreatePictureExample(wb);

            // === 圖表範例 ===
            // Program.CreateChartExample(wb, Program.testData);

            // === 匯出範例 ===
            // Program.CreateHtmlExportExample(wb);
            // Program.CreatePdfExportExample(wb);

            // === 進階範例 (注意：這些可能不需要 wb 參數或會建立新檔案，請小心使用) ===
            // Program.CreateSmartPipelineExample(Program.testData);
            // Program.CreateDomEditExample();
            // Program.ReadExcelExamples(wb);


        },
        session =>
        {
            // 使用 Shadow Copy (預設)，避免檔案鎖定問題
            // session.LibreOfficeOptions.UseShadowCopy = false;
        });
    }


}
