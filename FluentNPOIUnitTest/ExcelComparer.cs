using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace FluentNPOIUnitTest
{
    public static class ExcelComparer
    {
        public static string Compare(string path1, string path2, string sheetFilter = null)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"Comparing:\n  A: {path1}\n  B: {path2}");

            using var fs1 = new FileStream(path1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var fs2 = new FileStream(path2, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            IWorkbook wb1 = new XSSFWorkbook(fs1);
            IWorkbook wb2 = new XSSFWorkbook(fs2);

            if (wb1.NumberOfSheets != wb2.NumberOfSheets && string.IsNullOrEmpty(sheetFilter))
            {
                sb.AppendLine($"[SHEET COUNT MISMATCH] A: {wb1.NumberOfSheets}, B: {wb2.NumberOfSheets}");
            }

            for (int i = 0; i < Math.Min(wb1.NumberOfSheets, wb2.NumberOfSheets); i++)
            {
                var s1 = wb1.GetSheetAt(i);
                var s2 = wb2.GetSheetAt(i);

                if (!string.IsNullOrEmpty(sheetFilter) && s1.SheetName != sheetFilter && s2.SheetName != sheetFilter)
                    continue;

                CompareSheets(s1, s2, sb);
            }

            return sb.ToString();
        }

        private static void CompareSheets(ISheet s1, ISheet s2, StringBuilder sb)
        {
            sb.AppendLine($"\n--- Sheet: {s1.SheetName} vs {s2.SheetName} ---");

            if (s1.SheetName != s2.SheetName)
                sb.AppendLine($"  [NAME] '{s1.SheetName}' != '{s2.SheetName}'");

            // Pictures
            var pics1 = s1.Workbook.GetAllPictures().Count; // This gets ALL workbook pictures, but let's just check sheet drawings if possible.
            // Note: NPOI XSSF Sheet drawings are complex. For now, let's just check total rows.

            if (s1.PhysicalNumberOfRows != s2.PhysicalNumberOfRows)
                sb.AppendLine($"  [ROW COUNT] A: {s1.PhysicalNumberOfRows}, B: {s2.PhysicalNumberOfRows}");

            for (int r = 0; r <= Math.Max(s1.LastRowNum, s2.LastRowNum); r++)
            {
                var r1 = s1.GetRow(r);
                var r2 = s2.GetRow(r);

                if (r1 == null && r2 == null) continue;
                if (r1 == null || r2 == null)
                {
                    sb.AppendLine($"  [ROW {r}] One is null. A:{(r1 == null ? "null" : "ok")} B:{(r2 == null ? "null" : "ok")}");
                    continue;
                }

                if (r1.Height != r2.Height && Math.Abs(r1.Height - r2.Height) > 20) // Tolerance
                    sb.AppendLine($"  [ROW {r} HEIGHT] A: {r1.Height}, B: {r2.Height}");

                for (int c = 0; c < Math.Max(r1.LastCellNum, r2.LastCellNum); c++)
                {
                    var c1 = r1.GetCell(c);
                    var c2 = r2.GetCell(c);

                    CompareCells(c1, c2, r, c, s1.Workbook, s2.Workbook, sb);
                }
            }

            // Merged Regions
            CompareMergedRegions(s1, s2, sb);

            // Column Widths
            for (int c = 0; c < 20; c++) // Check first 20 cols
            {
                var w1 = s1.GetColumnWidth(c);
                var w2 = s2.GetColumnWidth(c);
                if (Math.Abs(w1 - w2) > 256) // Tolerance 1 char
                    sb.AppendLine($"  [COL {c} WIDTH] A: {w1}, B: {w2}");
            }
        }

        private static void CompareMergedRegions(ISheet s1, ISheet s2, StringBuilder sb)
        {
            if (s1.NumMergedRegions != s2.NumMergedRegions)
                sb.AppendLine($"  [MERGE COUNT] A: {s1.NumMergedRegions}, B: {s2.NumMergedRegions}");

            // Simple count check for now. Matching regions exactly is harder without sorting them.
        }

        private static void CompareCells(ICell c1, ICell c2, int r, int c, IWorkbook wb1, IWorkbook wb2, StringBuilder sb)
        {
            string prefix = $"  [CELL {r},{c}]";

            if (c1 == null && c2 == null) return;
            if (c1 == null || c2 == null)
            {
                sb.AppendLine($"{prefix} One is null. A:{(c1 == null ? "null" : "ok")} B:{(c2 == null ? "null" : "ok")}");
                return;
            }

            if (c1.CellType != c2.CellType)
                sb.AppendLine($"{prefix} TYPE A:{c1.CellType} B:{c2.CellType}");

            // Value
            string v1 = GetCellValue(c1);
            string v2 = GetCellValue(c2);
            if (v1 != v2)
                sb.AppendLine($"{prefix} VALUE A:'{v1}' B:'{v2}'");

            // Style
            CompareStyles(c1.CellStyle, c2.CellStyle, wb1, wb2, prefix, sb);
        }

        private static void CompareStyles(ICellStyle st1, ICellStyle st2, IWorkbook wb1, IWorkbook wb2, string prefix, StringBuilder sb)
        {
            if (st1 == null && st2 == null) return;
            if (st1 == null || st2 == null) return; // Already handled?

            // Font
            var f1 = st1.GetFont(wb1);
            var f2 = st2.GetFont(wb2);

            if (f1.FontName != f2.FontName) sb.AppendLine($"{prefix} FONT NAME A:{f1.FontName} B:{f2.FontName}");
            if (f1.FontHeightInPoints != f2.FontHeightInPoints) sb.AppendLine($"{prefix} FONT SIZE A:{f1.FontHeightInPoints} B:{f2.FontHeightInPoints}");
            if (f1.IsBold != f2.IsBold) sb.AppendLine($"{prefix} FONT BOLD A:{f1.IsBold} B:{f2.IsBold}");
            if (f1.Color != f2.Color) sb.AppendLine($"{prefix} FONT COLOR A:{f1.Color} B:{f2.Color}");

            // Fill
            if (st1.FillForegroundColor != st2.FillForegroundColor) sb.AppendLine($"{prefix} FILL FG A:{st1.FillForegroundColor} B:{st2.FillForegroundColor}");
            if (st1.FillPattern != st2.FillPattern) sb.AppendLine($"{prefix} FILL PATTERN A:{st1.FillPattern} B:{st2.FillPattern}");

            // Border
            if (st1.BorderTop != st2.BorderTop) sb.AppendLine($"{prefix} BORDER TOP A:{st1.BorderTop} B:{st2.BorderTop}");
            if (st1.BorderBottom != st2.BorderBottom) sb.AppendLine($"{prefix} BORDER BTM A:{st1.BorderBottom} B:{st2.BorderBottom}");

            // Alignment
            if (st1.Alignment != st2.Alignment) sb.AppendLine($"{prefix} ALIGN A:{st1.Alignment} B:{st2.Alignment}");
            if (st1.VerticalAlignment != st2.VerticalAlignment) sb.AppendLine($"{prefix} VALIGN A:{st1.VerticalAlignment} B:{st2.VerticalAlignment}");
        }

        private static string GetCellValue(ICell cell)
        {
            try
            {
                if (cell.CellType == CellType.Formula) return "=" + cell.CellFormula;
                return cell.ToString();
            }
            catch { return "Error"; }
        }
    }
}
