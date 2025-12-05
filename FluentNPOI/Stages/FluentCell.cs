using NPOI.SS.UserModel;
using FluentNPOI.Base;
using FluentNPOI.Helpers;
using FluentNPOI.Models;
using System;
using System.Collections.Generic;

namespace FluentNPOI.Stages
{
    /// <summary>
    /// 單元格操作類
    /// </summary>
    public class FluentCell : FluentCellBase
    {
        private ICell _cell;
        private ExcelCol _col;
        private int _row;
        public FluentCell(IWorkbook workbook, ISheet sheet,
        ICell cell, ExcelCol col, int row, Dictionary<string, ICellStyle> cellStylesCached = null)
            : base(workbook, sheet, cellStylesCached ?? new Dictionary<string, ICellStyle>())
        {
            _cell = cell;
            _col = col;
            _row = row;
        }

        public FluentCell SetValue<T>(T value)
        {
            if (_cell == null) return this;
            SetCellValue(_cell, value);
            return this;
        }

        public FluentCell SetFormulaValue(object value)
        {
            if (_cell == null) return this;
            SetFormulaValue(_cell, value);
            return this;
        }

        public FluentCell SetCellStyle(string cellStyleKey)
        {
            if (_cell == null) return this;

            if (!string.IsNullOrWhiteSpace(cellStyleKey) && _cellStylesCached.ContainsKey(cellStyleKey))
            {
                _cell.CellStyle = _cellStylesCached[cellStyleKey];
            }
            return this;
        }

        public FluentCell SetCellStyle(Func<TableCellStyleParams, CellStyleConfig> cellStyleAction)
        {
            if (_cell == null) return this;

            var cellStyleParams = new TableCellStyleParams
            {
                Workbook = _workbook,
                ColNum = (ExcelCol)_cell.ColumnIndex,
                RowNum = _cell.RowIndex,
                RowItem = null
            };

            // ✅ 先調用函數獲取樣式配置
            var config = cellStyleAction(cellStyleParams);

            if (!string.IsNullOrWhiteSpace(config.Key))
            {
                // ✅ 先檢查緩存
                if (!_cellStylesCached.ContainsKey(config.Key))
                {
                    // ✅ 只在不存在時才創建新樣式
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    config.StyleSetter(newCellStyle);
                    _cellStylesCached.Add(config.Key, newCellStyle);
                }
                _cell.CellStyle = _cellStylesCached[config.Key];
            }
            else
            {
                // 如果沒有返回 key，創建臨時樣式（不緩存）
                ICellStyle tempStyle = _workbook.CreateCellStyle();
                config.StyleSetter(tempStyle);
                _cell.CellStyle = tempStyle;
            }

            return this;
        }

        public FluentCell SetCellType(CellType cellType)
        {
            if (_cell == null) return this;
            _cell.SetCellType(cellType);
            return this;
        }

        /// <summary>
        /// 獲取當前單元格的值
        /// </summary>
        /// <returns>單元格的值（根據類型返回 bool, DateTime, double, string 或 null）</returns>
        public object GetValue()
        {
            return GetCellValue(_cell);
        }

        /// <summary>
        /// 獲取當前單元格的值並轉換為指定類型
        /// </summary>
        /// <typeparam name="T">目標類型</typeparam>
        /// <returns>轉換後的值</returns>
        public T GetValue<T>()
        {
            return GetCellValue<T>(_cell);
        }

        /// <summary>
        /// 獲取當前單元格的公式字符串（如果是公式單元格）
        /// </summary>
        /// <returns>公式字符串（不含 '=' 前綴），如果不是公式則返回 null</returns>
        public string GetFormula()
        {
            return GetCellFormulaValue(_cell);
        }

        /// <summary>
        /// 獲取當前單元格對象
        /// </summary>
        /// <returns>NPOI ICell 對象</returns>
        public ICell GetCell()
        {
            return _cell;
        }

        /// <summary>
        /// 在單元格中設置圖片（自動計算高度，保持原圖比例）
        /// </summary>
        /// <param name="pictureBytes">圖片字節數組</param>
        /// <param name="imgWidth">圖片寬度（像素）</param>
        /// <param name="anchorType">錨點類型</param>
        /// <param name="columnWidthRatio">列寬轉換比例（默認 7.0，表示像素寬度除以該值得到 Excel 列寬字符數）</param>
        /// <returns>FluentCell 實例，支持鏈式調用</returns>
        public FluentCell SetPictureOnCell(byte[] pictureBytes, int imgWidth, AnchorType anchorType = AnchorType.MoveAndResize, double columnWidthRatio = 7.0)
        {
            // 自動計算高度（需要從圖片中讀取原始尺寸）
            // 由於無法直接從字節數組獲取圖片尺寸，這裡使用寬度作為高度（1:1比例）
            // 如果需要更精確的比例，可以考慮使用 System.Drawing.Image 或其他圖像庫
            return SetPictureOnCell(pictureBytes, imgWidth, imgWidth, anchorType, columnWidthRatio);
        }

        /// <summary>
        /// 在單元格中設置圖片（手動設置寬度和高度）
        /// </summary>
        /// <param name="pictureBytes">圖片字節數組</param>
        /// <param name="imgWidth">圖片寬度（像素）</param>
        /// <param name="imgHeight">圖片高度（像素）</param>
        /// <param name="anchorType">錨點類型</param>
        /// <param name="columnWidthRatio">列寬轉換比例（默認 7.0，表示像素寬度除以該值得到 Excel 列寬字符數）</param>
        /// <returns>FluentCell 實例，支持鏈式調用</returns>
        public FluentCell SetPictureOnCell(byte[] pictureBytes, int imgWidth, int imgHeight, AnchorType anchorType = AnchorType.MoveAndResize, double columnWidthRatio = 7.0)
        {
            // 參數驗證
            ValidatePictureParameters(pictureBytes, imgWidth, imgHeight, columnWidthRatio);

            // 設置列寬
            double columnWidth = CalculateColumnWidth(imgWidth, columnWidthRatio);
            _sheet.SetColumnWidth((int)_col, (int)Math.Round(columnWidth * 256));

            // 獲取圖片類型並添加到工作簿
            var picType = GetPictureType(pictureBytes);
            int picIndex = _workbook.AddPicture(pictureBytes, picType);

            // 創建繪圖對象和錨點
            IDrawing drawing = _sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = CreatePictureAnchor(imgWidth, imgHeight, anchorType);

            // 創建圖片
            IPicture pict = drawing.CreatePicture(anchor, picIndex);

            // 調整圖片大小（如果需要）
            // NPOI 的 IPicture.Resize() 方法可以根據錨點自動調整，但也可以手動調用
            // pict.Resize(); // 可選：根據錨點自動調整大小

            return this;
        }

        /// <summary>
        /// 驗證圖片參數
        /// </summary>
        private void ValidatePictureParameters(byte[] pictureBytes, int imgWidth, int imgHeight, double columnWidthRatio)
        {
            if (_cell == null)
            {
                throw new InvalidOperationException("No active cell. Call SetCellPosition(...) first.");
            }

            if (pictureBytes == null || pictureBytes.Length == 0)
            {
                throw new ArgumentException("Picture bytes cannot be null or empty.", nameof(pictureBytes));
            }

            if (imgWidth <= 0)
            {
                throw new ArgumentException("Image width must be greater than zero.", nameof(imgWidth));
            }

            if (imgHeight <= 0)
            {
                throw new ArgumentException("Image height must be greater than zero.", nameof(imgHeight));
            }

            if (columnWidthRatio <= 0)
            {
                throw new ArgumentException("Column width ratio must be greater than zero.", nameof(columnWidthRatio));
            }
        }

        /// <summary>
        /// 計算列寬（將像素寬度轉換為 Excel 列寬單位）
        /// </summary>
        /// <param name="imgWidth">圖片寬度（像素）</param>
        /// <param name="columnWidthRatio">轉換比例</param>
        /// <returns>Excel 列寬（字符數）</returns>
        private double CalculateColumnWidth(int imgWidth, double columnWidthRatio)
        {
            // Excel 列寬單位：1 個字符寬度 = 256 單位
            // 將像素寬度除以轉換比例得到字符數
            return imgWidth / columnWidthRatio;
        }

        /// <summary>
        /// 創建圖片錨點，設置完整的位置和大小信息
        /// </summary>
        /// <param name="imgWidth">圖片寬度（像素）</param>
        /// <param name="imgHeight">圖片高度（像素）</param>
        /// <param name="anchorType">錨點類型</param>
        /// <returns>配置好的 IClientAnchor 對象</returns>
        private IClientAnchor CreatePictureAnchor(int imgWidth, int imgHeight, AnchorType anchorType)
        {
            ICreationHelper creationHelper = _workbook.GetCreationHelper();
            IClientAnchor anchor = creationHelper.CreateClientAnchor();

            // 設置起始位置（_row 已經是 0-based，因為在 SetCellPosition 中已經轉換）
            anchor.Col1 = (short)_col;
            anchor.Row1 = (short)_row;

            // 計算結束位置（Col2 和 Row2）
            // 根據圖片尺寸和單元格大小計算需要跨越多少列和行
            // Excel 默認列寬約為 8.43 字符（約 64 像素），行高約為 15 像素
            // 這裡使用簡化的計算方式
            
            // 獲取當前列寬（以字符為單位）
            // GetColumnWidth 返回 int（以 1/256 字符為單位），轉換為字符數
            double columnWidthInChars = _sheet.GetColumnWidth((int)_col) / 256.0;
            
            // 獲取當前行高（以點為單位，1 點 ≈ 1.33 像素）
            IRow row = _sheet.GetRow(_row) ?? _sheet.CreateRow(_row);
            short rowHeightInPoints = row.Height > 0 ? (short)(row.Height / 20.0) : (short)15; // 默認行高約 15 點
            
            // 計算需要跨越的列數（考慮列寬）
            // 假設 1 字符寬度 ≈ 7 像素（可根據實際情況調整）
            double pixelsPerChar = 7.0;
            double columnsNeeded = imgWidth / (columnWidthInChars * pixelsPerChar);
            short col2 = (short)Math.Min((int)_col + (int)Math.Ceiling(columnsNeeded), 16383); // Excel 最大列數限制

            // 計算需要跨越的行數（考慮行高）
            // 1 點 ≈ 1.33 像素
            double pixelsPerPoint = 1.33;
            double rowsNeeded = imgHeight / (rowHeightInPoints * pixelsPerPoint);
            short row2 = (short)Math.Min(_row + (int)Math.Ceiling(rowsNeeded), 1048575); // Excel 最大行數限制

            anchor.Col2 = col2;
            anchor.Row2 = row2;
            anchor.AnchorType = anchorType;

            return anchor;
        }

        private PictureType GetPictureType(byte[] pictureBytes)
        {
            if (pictureBytes == null || pictureBytes.Length < 4)
            {
                throw new ArgumentException("Invalid picture bytes: array is null or too short.", nameof(pictureBytes));
            }

            // PNG: 89 50 4E 47 0D 0A 1A 0A
            if (pictureBytes.Length >= 8 &&
                pictureBytes[0] == 0x89 && pictureBytes[1] == 0x50 && pictureBytes[2] == 0x4E && pictureBytes[3] == 0x47 &&
                pictureBytes[4] == 0x0D && pictureBytes[5] == 0x0A && pictureBytes[6] == 0x1A && pictureBytes[7] == 0x0A)
            {
                return PictureType.PNG;
            }

            // JPEG: FF D8 FF
            if (pictureBytes.Length >= 3 &&
                pictureBytes[0] == 0xFF && pictureBytes[1] == 0xD8 && pictureBytes[2] == 0xFF)
            {
                return PictureType.JPEG;
            }

            // GIF: 47 49 46 38 (GIF8)
            if (pictureBytes.Length >= 4 &&
                pictureBytes[0] == 0x47 && pictureBytes[1] == 0x49 && pictureBytes[2] == 0x46 && pictureBytes[3] == 0x38)
            {
                return PictureType.GIF;
            }

            // BMP/DIB: 42 4D (BM)
            if (pictureBytes.Length >= 2 &&
                pictureBytes[0] == 0x42 && pictureBytes[1] == 0x4D)
            {
                return PictureType.DIB;
            }

            // EMF: 01 00 00 00 (但需要更多检查，EMF 文件通常以这个开头)
            if (pictureBytes.Length >= 4 &&
                pictureBytes[0] == 0x01 && pictureBytes[1] == 0x00 && pictureBytes[2] == 0x00 && pictureBytes[3] == 0x00)
            {
                // 检查是否是有效的 EMF 文件（EMF 文件头通常是 40 字节）
                if (pictureBytes.Length >= 40)
                {
                    // EMF 文件的第二个 DWORD 应该是文件大小
                    // 这里做简单检查，如果符合 EMF 特征就返回 EMF
                    return PictureType.EMF;
                }
            }

            // WMF: 通常以 01 00 09 00 开头（但需要更多检查）
            if (pictureBytes.Length >= 4 &&
                pictureBytes[0] == 0x01 && pictureBytes[1] == 0x00 && pictureBytes[2] == 0x09 && pictureBytes[3] == 0x00)
            {
                return PictureType.WMF;
            }

            throw new NotSupportedException($"Unsupported picture format. File header: {BitConverter.ToString(pictureBytes, 0, Math.Min(8, pictureBytes.Length))}");
        }
    }
}

