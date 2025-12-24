using NPOI.SS.UserModel;

namespace FluentNPOI.HotReload.Styling;

/// <summary>
/// Manages cell styles with caching to comply with Excel's 4,000 unique style limit.
/// Automatically deduplicates styles across the widget tree during hot reload cycles.
/// </summary>
public class StyleManager
{
    private readonly Dictionary<string, ICellStyle> _styleCache = new();
    private IWorkbook? _workbook;

    /// <summary>
    /// Gets the current number of cached styles.
    /// </summary>
    public int CachedStyleCount => _styleCache.Count;

    /// <summary>
    /// Resets the style cache for a new workbook instance.
    /// Should be called at the start of each hot reload refresh.
    /// </summary>
    /// <param name="workbook">The new workbook instance.</param>
    public void Reset(IWorkbook workbook)
    {
        _styleCache.Clear();
        _workbook = workbook;
    }

    /// <summary>
    /// Gets or creates a cached cell style based on the FluentStyle definition.
    /// </summary>
    /// <param name="style">The FluentStyle definition.</param>
    /// <returns>The NPOI cell style.</returns>
    public ICellStyle GetOrCreateStyle(FluentStyle style)
    {
        if (_workbook == null)
            throw new InvalidOperationException("StyleManager not initialized. Call Reset() first.");

        var cacheKey = style.GetCacheKey();

        if (_styleCache.TryGetValue(cacheKey, out var existingStyle))
            return existingStyle;

        var cellStyle = CreateStyle(style);
        _styleCache[cacheKey] = cellStyle;
        return cellStyle;
    }

    /// <summary>
    /// Gets or creates a cached cell style and applies it to the specified cell.
    /// </summary>
    /// <param name="cell">The cell to apply the style to.</param>
    /// <param name="style">The FluentStyle definition.</param>
    public void ApplyStyle(ICell cell, FluentStyle style)
    {
        if (!style.HasAnyStyle())
            return;

        var cellStyle = GetOrCreateStyle(style);
        cell.CellStyle = cellStyle;
    }

    private ICellStyle CreateStyle(FluentStyle style)
    {
        var cellStyle = _workbook!.CreateCellStyle();

        // Background color
        if (style.BackgroundColor != null)
        {
            cellStyle.FillPattern = FillPattern.SolidForeground;
            cellStyle.FillForegroundColor = style.BackgroundColor.Index;
        }

        // Border
        if (style.Border.HasValue)
        {
            cellStyle.BorderTop = style.Border.Value;
            cellStyle.BorderBottom = style.Border.Value;
            cellStyle.BorderLeft = style.Border.Value;
            cellStyle.BorderRight = style.Border.Value;
        }

        // Alignment
        if (style.HorizontalAlignment.HasValue)
            cellStyle.Alignment = style.HorizontalAlignment.Value;

        if (style.VerticalAlignment.HasValue)
            cellStyle.VerticalAlignment = style.VerticalAlignment.Value;

        // Wrap text
        if (style.WrapText.HasValue)
            cellStyle.WrapText = style.WrapText.Value;

        // Number format
        if (!string.IsNullOrEmpty(style.NumberFormat))
        {
            var format = _workbook.CreateDataFormat();
            cellStyle.DataFormat = format.GetFormat(style.NumberFormat);
        }

        // Font
        if (style.IsBold || style.IsItalic || style.FontColor != null ||
            style.FontName != null || style.FontSize.HasValue)
        {
            var font = _workbook.CreateFont();

            if (style.IsBold)
                font.IsBold = true;

            if (style.IsItalic)
                font.IsItalic = true;

            if (style.FontColor != null)
                font.Color = style.FontColor.Index;

            if (style.FontName != null)
                font.FontName = style.FontName;

            if (style.FontSize.HasValue)
                font.FontHeightInPoints = style.FontSize.Value;

            cellStyle.SetFont(font);
        }

        return cellStyle;
    }

    /// <summary>
    /// Gets the default style with no modifications.
    /// </summary>
    public ICellStyle GetDefaultStyle()
    {
        if (_workbook == null)
            throw new InvalidOperationException("StyleManager not initialized. Call Reset() first.");

        const string defaultKey = "__default__";
        if (_styleCache.TryGetValue(defaultKey, out var defaultStyle))
            return defaultStyle;

        var style = _workbook.CreateCellStyle();
        _styleCache[defaultKey] = style;
        return style;
    }
}
