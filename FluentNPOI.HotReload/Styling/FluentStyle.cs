using NPOI.SS.UserModel;

namespace FluentNPOI.HotReload.Styling;

/// <summary>
/// Fluent style builder for declarative cell styling.
/// Generates cache keys for style deduplication to comply with Excel's 4,000 style limit.
/// </summary>
public class FluentStyle
{
    /// <summary>
    /// Gets or sets the font color.
    /// </summary>
    public IndexedColors? FontColor { get; private set; }

    /// <summary>
    /// Gets or sets the background color.
    /// </summary>
    public IndexedColors? BackgroundColor { get; private set; }

    /// <summary>
    /// Gets or sets whether the text is bold.
    /// </summary>
    public bool IsBold { get; private set; }

    /// <summary>
    /// Gets or sets whether the text is italic.
    /// </summary>
    public bool IsItalic { get; private set; }

    /// <summary>
    /// Gets or sets the border style.
    /// </summary>
    public BorderStyle? Border { get; private set; }

    /// <summary>
    /// Gets or sets the font name.
    /// </summary>
    public string? FontName { get; private set; }

    /// <summary>
    /// Gets or sets the font size in points.
    /// </summary>
    public double? FontSize { get; private set; }

    /// <summary>
    /// Gets or sets the horizontal alignment.
    /// </summary>
    public HorizontalAlignment? HorizontalAlignment { get; private set; }

    /// <summary>
    /// Gets or sets the vertical alignment.
    /// </summary>
    public VerticalAlignment? VerticalAlignment { get; private set; }

    /// <summary>
    /// Gets or sets the number format string.
    /// </summary>
    public string? NumberFormat { get; private set; }

    /// <summary>
    /// Gets or sets whether text wrapping is enabled.
    /// </summary>
    public bool? WrapText { get; private set; }

    /// <summary>
    /// Sets the font color.
    /// </summary>
    public FluentStyle SetFontColor(IndexedColors color)
    {
        FontColor = color;
        return this;
    }

    /// <summary>
    /// Sets the background color.
    /// </summary>
    public FluentStyle SetBackgroundColor(IndexedColors color)
    {
        BackgroundColor = color;
        return this;
    }

    /// <summary>
    /// Sets the text to bold.
    /// </summary>
    public FluentStyle SetBold(bool bold = true)
    {
        IsBold = bold;
        return this;
    }

    /// <summary>
    /// Sets the text to italic.
    /// </summary>
    public FluentStyle SetItalic(bool italic = true)
    {
        IsItalic = italic;
        return this;
    }

    /// <summary>
    /// Sets the border style for all sides.
    /// </summary>
    public FluentStyle SetBorder(BorderStyle style)
    {
        Border = style;
        return this;
    }

    /// <summary>
    /// Sets the font name.
    /// </summary>
    public FluentStyle SetFontName(string fontName)
    {
        FontName = fontName;
        return this;
    }

    /// <summary>
    /// Sets the font size in points.
    /// </summary>
    public FluentStyle SetFontSize(double size)
    {
        FontSize = size;
        return this;
    }

    /// <summary>
    /// Sets the horizontal alignment.
    /// </summary>
    public FluentStyle SetHorizontalAlignment(HorizontalAlignment alignment)
    {
        HorizontalAlignment = alignment;
        return this;
    }

    /// <summary>
    /// Sets the vertical alignment.
    /// </summary>
    public FluentStyle SetVerticalAlignment(VerticalAlignment alignment)
    {
        VerticalAlignment = alignment;
        return this;
    }

    /// <summary>
    /// Sets the number format.
    /// </summary>
    public FluentStyle SetNumberFormat(string format)
    {
        NumberFormat = format;
        return this;
    }

    /// <summary>
    /// Sets whether text should wrap.
    /// </summary>
    public FluentStyle SetWrapText(bool wrap = true)
    {
        WrapText = wrap;
        return this;
    }

    /// <summary>
    /// Generates a unique cache key for style deduplication.
    /// </summary>
    public string GetCacheKey()
    {
        return string.Join("|",
            FontColor?.Index.ToString() ?? "null",
            BackgroundColor?.Index.ToString() ?? "null",
            IsBold.ToString(),
            IsItalic.ToString(),
            Border?.ToString() ?? "null",
            FontName ?? "null",
            FontSize?.ToString() ?? "null",
            HorizontalAlignment?.ToString() ?? "null",
            VerticalAlignment?.ToString() ?? "null",
            NumberFormat ?? "null",
            WrapText?.ToString() ?? "null"
        );
    }

    /// <summary>
    /// Checks if the style has any properties set.
    /// </summary>
    public bool HasAnyStyle()
    {
        return FontColor != null ||
               BackgroundColor != null ||
               IsBold ||
               IsItalic ||
               Border != null ||
               FontName != null ||
               FontSize != null ||
               HorizontalAlignment != null ||
               VerticalAlignment != null ||
               NumberFormat != null ||
               WrapText != null;
    }
}
