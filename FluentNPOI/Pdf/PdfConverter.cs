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

                // Render rows with merge support
                for (int r = 0; r <= sheet.LastRowNum; r++)
                {
                    var row = sheet.GetRow(r);

                    for (int c = 0; c < maxCol; c++)
                    {
                        // Check if inside merged region
                        var region = GetMergedRegion(sheet, r, c);

                        // Skip cells that are inside a merged region but not the first cell
                        if (region != null && (r != region.FirstRow || c != region.FirstColumn))
                        {
                            continue;
                        }

                        ICell cell = row?.GetCell(c);
                        var cellStyle = cell?.CellStyle;

                        // Calculate span
                        uint rowSpan = 1;
                        uint colSpan = 1;
                        if (region != null)
                        {
                            rowSpan = (uint)(region.LastRow - region.FirstRow + 1);
                            colSpan = (uint)(region.LastColumn - region.FirstColumn + 1);
                        }

                        // Create cell with explicit position and span
                        var tableCell = table.Cell()
                            .Row((uint)(r + 1))
                            .Column((uint)(c + 1))
                            .RowSpan(rowSpan)
                            .ColumnSpan(colSpan)
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

                        // Text content with alignment
                        var text = GetCellText(cell);

                        // Apply alignment
                        IContainer alignedCell = tableCell;
                        if (cellStyle != null)
                        {
                            switch (cellStyle.Alignment)
                            {
                                case NPOI.SS.UserModel.HorizontalAlignment.Center:
                                case NPOI.SS.UserModel.HorizontalAlignment.CenterSelection:
                                    alignedCell = tableCell.AlignCenter();
                                    break;
                                case NPOI.SS.UserModel.HorizontalAlignment.Right:
                                    alignedCell = tableCell.AlignRight();
                                    break;
                                default:
                                    alignedCell = tableCell.AlignLeft();
                                    break;
                            }
                        }

                        if (cellStyle != null)
                        {
                            var font = cellStyle.GetFont(workbook);
                            if (font != null)
                            {
                                var textBlock = alignedCell.Text(text);
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
                                alignedCell.Text(text);
                            }
                        }
                        else
                        {
                            alignedCell.Text(text);
                        }
                    }
                }
            });
        }

        private static CellRangeAddress GetMergedRegion(ISheet sheet, int row, int col)
        {
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                var region = sheet.GetMergedRegion(i);
                if (row >= region.FirstRow && row <= region.LastRow &&
                    col >= region.FirstColumn && col <= region.LastColumn)
                {
                    return region;
                }
            }
            return null;
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

            // Use NPOI DataFormatter for proper formatting
            var formatter = new NPOI.SS.UserModel.DataFormatter();
            try
            {
                return formatter.FormatCellValue(cell);
            }
            catch
            {
                // Fallback for formula cells or errors
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue ?? "";
                    case CellType.Numeric:
                        return cell.NumericCellValue.ToString();
                    case CellType.Boolean:
                        return cell.BooleanCellValue ? "TRUE" : "FALSE";
                    default:
                        return "";
                }
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
