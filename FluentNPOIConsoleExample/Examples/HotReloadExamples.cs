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

        FluentLivePreview.Run("output/demo.xlsx", wb =>
        {
            // Set up necessary styles first
            Program.SetupStyles(wb);

            // === Basic Examples ===
            Program.CreateBasicTableExample(wb, Program.testData);
            Program.CreateSummaryExample(wb, Program.testData);
            Program.CreateDataTableExample(wb);

            // === Style Examples ===
            // Program.CreateCellStyleRangeExample(wb);
            // Program.CreateCopyStyleExample(wb, Program.testData); // (Part 2)
            // Program.CreateSheetGlobalStyleExample(wb, Program.testData); // (Part 3)
            // Program.CreateMappingStylingExample(wb, Program.testData); // (New Feature)

            // === Cell Operations ===
            // Program.CreateSetCellValueExample(wb);
            // Program.CreateCellMergeExample(wb);
            // Program.CreatePictureExample(wb);

            // === Chart Examples ===
            // Program.CreateChartExample(wb, Program.testData);

            // === Export Examples ===
            // Program.CreateHtmlExportExample(wb);
            // Program.CreatePdfExportExample(wb);

            // === Advanced Examples (Note: These might not need 'wb' or may create new files) ===
            // Program.CreateSmartPipelineExample(Program.testData);
            // Program.CreateDomEditExample();
            // Program.ReadExcelExamples(wb);


        },
        session =>
        {
            // Configure session here if needed
        });

        // Example: Reading from an existing template
        // FluentLivePreview.RunFromTemplate("template.xlsx", "output/modified_template.xlsx", wb =>
        // {
        //     wb.UseSheet("Sheet1")
        //       .SetCellPosition(ExcelCol.A, 1)
        //       .SetValue("Modified!");
        // });

        // Example: Reading from a Stream
        // using var fs = File.OpenRead("template.xlsx");
        // FluentLivePreview.RunFromTemplate(fs, "output/stream_modified.xlsx", wb =>
        // {
        //     wb.UseSheet("Sheet1")
        //       .SetCellPosition(ExcelCol.B, 2)
        //       .SetValue("From Stream!");
        // });
    }


}
