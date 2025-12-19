using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace FluentNPOI.Pdf
{
    /// <summary>
    /// Helper class to convert Excel Sheet to PDF using QuestPDF
    /// </summary>
    public static class PdfConverter
    {
        /// <summary>
        /// Convert an ISheet to PDF and save to file
        /// </summary>
        public static void ConvertSheetToPdf(ISheet sheet, IWorkbook workbook, string outputPath)
        {
            // Enable QuestPDF Community License
            QuestPDF.Settings.License = LicenseType.Community;

            try
            {
                Document.Create(container =>
                {
                    container.Page(page =>
                    {
                        page.Size(PageSizes.A4.Landscape());
                        page.Margin(1, Unit.Centimetre);
                        page.DefaultTextStyle(x => x.FontSize(10));

                        page.Content().Element(c => BuildTable(c, sheet, workbook));
                    });
                }).GeneratePdf(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PDF Generation Error: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Convert an ISheet to PDF bytes
        /// </summary>
        public static byte[] ConvertSheetToPdfBytes(ISheet sheet, IWorkbook workbook)
        {
            QuestPDF.Settings.License = LicenseType.Community;

            return Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(1, Unit.Centimetre);
                    page.DefaultTextStyle(x => x.FontSize(10));

                    page.Content().Element(c => BuildTable(c, sheet, workbook));
                });
            }).GeneratePdf();
        }

        private static void BuildTable(IContainer container, ISheet sheet, IWorkbook workbook)
        {
            // Build merge map
            var mergeMap = BuildMergeMap(sheet);

            // Calculate column count
            int maxCol = 0;
            for (int r = 0; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row != null && row.LastCellNum > maxCol)
                    maxCol = row.LastCellNum;
            }

            if (maxCol == 0) return;

            container.Table(table =>
            {
                // Define columns
                table.ColumnsDefinition(columns =>
                {
                    for (int c = 0; c < maxCol; c++)
                    {
                        columns.RelativeColumn();
                    }
                });

                // Render rows - simplified version (no merge support for now)
                for (int r = 0; r <= sheet.LastRowNum; r++)
                {
                    var row = sheet.GetRow(r);

                    for (int c = 0; c < maxCol; c++)
                    {
                        ICell cell = row?.GetCell(c);
                        var cellStyle = cell?.CellStyle;

                        // Create cell with explicit position
                        var tableCell = table.Cell()
                            .Row((uint)(r + 1))
                            .Column((uint)(c + 1))
                            .Padding(3);

                        // Apply borders from Excel style
                        if (cellStyle != null)
                        {
                            ApplyBorders(ref tableCell, cellStyle, workbook);
                        }
                        else
                        {
                            tableCell = tableCell.Border(0.5f).BorderColor(Colors.Grey.Lighten1);
                        }

                        // Background color
                        if (cellStyle != null && cellStyle.FillPattern == FillPattern.SolidForeground)
                        {
                            var bgColor = GetQuestColor(cellStyle.FillForegroundColorColor, cellStyle.FillForegroundColor, workbook);
                            if (bgColor.HasValue)
                                tableCell = tableCell.Background(bgColor.Value);
                        }

                        // Text content
                        var text = GetCellText(cell);
                        if (cellStyle != null)
                        {
                            var font = cellStyle.GetFont(workbook);
                            if (font != null)
                            {
                                var textBlock = tableCell.Text(text);
                                if (font.IsBold)
                                    textBlock.Bold();
                                if (font.IsItalic)
                                    textBlock.Italic();
                                if (font.Underline != FontUnderlineType.None)
                                    textBlock.Underline();
                                if (font.IsStrikeout)
                                    textBlock.Strikethrough();
                                if (font.FontHeightInPoints > 0)
                                    textBlock.FontSize((float)font.FontHeightInPoints);

                                var fontColor = GetFontQuestColor(font, workbook);
                                if (fontColor.HasValue)
                                    textBlock.FontColor(fontColor.Value);
                            }
                            else
                            {
                                tableCell.Text(text);
                            }
                        }
                        else
                        {
                            tableCell.Text(text);
                        }
                    }
                }
            });
        }

        private static void ApplyBorders(ref IContainer container, ICellStyle style, IWorkbook workbook)
        {
            // Get border width and color for each side
            var topWidth = GetBorderWidth(style.BorderTop);
            var rightWidth = GetBorderWidth(style.BorderRight);
            var bottomWidth = GetBorderWidth(style.BorderBottom);
            var leftWidth = GetBorderWidth(style.BorderLeft);

            // If no borders, use default thin grey
            if (topWidth == 0 && rightWidth == 0 && bottomWidth == 0 && leftWidth == 0)
            {
                container = container.Border(0.5f).BorderColor(Colors.Grey.Lighten1);
                return;
            }

            // Apply each border separately
            if (topWidth > 0)
            {
                var color = GetColorFromIndex(style.TopBorderColor, workbook) ?? Color.FromRGB(0, 0, 0);
                container = container.BorderTop(topWidth).BorderColor(color);
            }
            if (rightWidth > 0)
            {
                var color = GetColorFromIndex(style.RightBorderColor, workbook) ?? Color.FromRGB(0, 0, 0);
                container = container.BorderRight(rightWidth).BorderColor(color);
            }
            if (bottomWidth > 0)
            {
                var color = GetColorFromIndex(style.BottomBorderColor, workbook) ?? Color.FromRGB(0, 0, 0);
                container = container.BorderBottom(bottomWidth).BorderColor(color);
            }
            if (leftWidth > 0)
            {
                var color = GetColorFromIndex(style.LeftBorderColor, workbook) ?? Color.FromRGB(0, 0, 0);
                container = container.BorderLeft(leftWidth).BorderColor(color);
            }
        }

        private static float GetBorderWidth(BorderStyle border)
        {
            switch (border)
            {
                case BorderStyle.Thin:
                case BorderStyle.Hair:
                case BorderStyle.Dashed:
                case BorderStyle.Dotted:
                    return 0.5f;
                case BorderStyle.Medium:
                case BorderStyle.MediumDashed:
                case BorderStyle.MediumDashDot:
                case BorderStyle.MediumDashDotDot:
                    return 1.5f;
                case BorderStyle.Thick:
                case BorderStyle.Double:
                    return 2f;
                default:
                    return 0f;
            }
        }

        private static string GetCellText(ICell cell)
        {
            if (cell == null) return "";

            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue ?? "";
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return string.Format("{0:yyyy-MM-dd}", cell.DateCellValue);
                    return cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue ? "TRUE" : "FALSE";
                case CellType.Formula:
                    try { return cell.NumericCellValue.ToString(); }
                    catch { return cell.StringCellValue ?? ""; }
                default:
                    return "";
            }
        }

        private static Color? GetQuestColor(IColor color, short colorIndex, IWorkbook workbook)
        {
            // Try XSSFColor first
            if (color is NPOI.XSSF.UserModel.XSSFColor xColor && xColor.IsRGB)
            {
                var rgb = xColor.RGB;
                if (rgb != null && rgb.Length >= 3)
                    return Color.FromRGB(rgb[0], rgb[1], rgb[2]);
            }

            // Fallback to indexed color
            return GetColorFromIndex(colorIndex, workbook);
        }

        private static Color? GetFontQuestColor(IFont font, IWorkbook workbook)
        {
            if (font is NPOI.XSSF.UserModel.XSSFFont xFont)
            {
                var xColor = xFont.GetXSSFColor();
                if (xColor != null && xColor.IsRGB)
                {
                    var rgb = xColor.RGB;
                    if (rgb != null && rgb.Length >= 3)
                        return Color.FromRGB(rgb[0], rgb[1], rgb[2]);
                }
            }

            return GetColorFromIndex(font.Color, workbook);
        }

        private static Color? GetColorFromIndex(short index, IWorkbook workbook)
        {
            if (index == IndexedColors.Automatic.Index || index == 0) return null;

            // Standard color map
            if (StandardColorMap.TryGetValue(index, out var color))
                return color;

            // HSSF Palette
            if (workbook is NPOI.HSSF.UserModel.HSSFWorkbook hVal)
            {
                var palette = hVal.GetCustomPalette();
                var hColor = palette.GetColor(index);
                if (hColor != null)
                {
                    var rgb = hColor.RGB;
                    return Color.FromRGB(rgb[0], rgb[1], rgb[2]);
                }
            }

            return null;
        }

        private static bool IsInsideMergedRegion(ISheet sheet, int row, int col, out CellRangeAddress region)
        {
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                var r = sheet.GetMergedRegion(i);
                if (row >= r.FirstRow && row <= r.LastRow &&
                    col >= r.FirstColumn && col <= r.LastColumn)
                {
                    region = r;
                    return true;
                }
            }
            region = null;
            return false;
        }

        private static Dictionary<(int, int), CellRangeAddress> BuildMergeMap(ISheet sheet)
        {
            var map = new Dictionary<(int, int), CellRangeAddress>();
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                var region = sheet.GetMergedRegion(i);
                map[(region.FirstRow, region.FirstColumn)] = region;
            }
            return map;
        }

        // Standard color map (same as HtmlConverter)
        private static readonly Dictionary<short, Color> StandardColorMap = new Dictionary<short, Color>
        {
            { IndexedColors.Black.Index, Color.FromRGB(0, 0, 0) },
            { IndexedColors.White.Index, Color.FromRGB(255, 255, 255) },
            { IndexedColors.Red.Index, Color.FromRGB(255, 0, 0) },
            { IndexedColors.BrightGreen.Index, Color.FromRGB(0, 255, 0) },
            { IndexedColors.Blue.Index, Color.FromRGB(0, 0, 255) },
            { IndexedColors.Yellow.Index, Color.FromRGB(255, 255, 0) },
            { IndexedColors.Pink.Index, Color.FromRGB(255, 0, 255) },
            { IndexedColors.Turquoise.Index, Color.FromRGB(0, 255, 255) },
            { IndexedColors.DarkRed.Index, Color.FromRGB(128, 0, 0) },
            { IndexedColors.Green.Index, Color.FromRGB(0, 128, 0) },
            { IndexedColors.DarkBlue.Index, Color.FromRGB(0, 0, 128) },
            { IndexedColors.DarkYellow.Index, Color.FromRGB(128, 128, 0) },
            { IndexedColors.Violet.Index, Color.FromRGB(128, 0, 128) },
            { IndexedColors.Teal.Index, Color.FromRGB(0, 128, 128) },
            { IndexedColors.Grey25Percent.Index, Color.FromRGB(192, 192, 192) },
            { IndexedColors.Grey50Percent.Index, Color.FromRGB(128, 128, 128) },
            { IndexedColors.LightCornflowerBlue.Index, Color.FromRGB(204, 204, 255) },
            { IndexedColors.Orange.Index, Color.FromRGB(255, 165, 0) },
            { IndexedColors.Gold.Index, Color.FromRGB(255, 215, 0) },
            { IndexedColors.Lime.Index, Color.FromRGB(0, 255, 0) },
            { IndexedColors.Aqua.Index, Color.FromRGB(0, 255, 255) },
        };
    }
}
