using System.Runtime.CompilerServices;
using FluentNPOI.HotReload.Context;

namespace FluentNPOI.HotReload.Widgets;

/// <summary>
/// A widget that arranges its children horizontally with automatic column width calculation
/// based on Weight properties. Similar to CSS Flexbox.
/// </summary>
public class FlexibleRow : ExcelWidget
{
    /// <summary>
    /// Gets the list of child widgets (cells).
    /// </summary>
    public List<ExcelWidget> Cells { get; }

    /// <summary>
    /// The default weight for children without explicit Weight property.
    /// </summary>
    public int DefaultWeight { get; set; } = 1;

    /// <summary>
    /// The total available width in Excel column units (characters).
    /// Default is 100 characters total.
    /// </summary>
    public int TotalWidth { get; set; } = 100;

    /// <summary>
    /// Minimum column width in characters.
    /// </summary>
    public int MinColumnWidth { get; set; } = 5;

    /// <summary>
    /// Creates a new FlexibleRow widget with the specified cells.
    /// </summary>
    /// <param name="cells">The child widgets to arrange horizontally.</param>
    /// <param name="filePath">Auto-captured source file path.</param>
    /// <param name="lineNumber">Auto-captured line number.</param>
    public FlexibleRow(
        IEnumerable<ExcelWidget> cells,
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0)
        : base(filePath, lineNumber)
    {
        Cells = cells.ToList();
    }

    /// <summary>
    /// Creates a new FlexibleRow widget with the specified cells.
    /// </summary>
    /// <param name="cells">The child widgets to arrange horizontally.</param>
    public FlexibleRow(params ExcelWidget[] cells)
        : this(cells.AsEnumerable())
    {
    }

    /// <summary>
    /// Sets the total available width and returns this instance for fluent chaining.
    /// </summary>
    /// <param name="width">The total width in characters.</param>
    /// <returns>This FlexibleRow instance.</returns>
    public FlexibleRow SetTotalWidth(int width)
    {
        TotalWidth = width;
        return this;
    }

    /// <summary>
    /// Sets the default weight for children without explicit Weight property.
    /// </summary>
    /// <param name="weight">The default weight value.</param>
    /// <returns>This FlexibleRow instance.</returns>
    public FlexibleRow SetDefaultWeight(int weight)
    {
        DefaultWeight = weight;
        return this;
    }

    /// <summary>
    /// Sets the minimum column width.
    /// </summary>
    /// <param name="width">The minimum width in characters.</param>
    /// <returns>This FlexibleRow instance.</returns>
    public FlexibleRow SetMinColumnWidth(int width)
    {
        MinColumnWidth = width;
        return this;
    }

    /// <inheritdoc/>
    public override void Build(ExcelContext ctx)
    {
        if (Cells.Count == 0)
            return;

        // Calculate total weight
        var totalWeight = Cells.Sum(c => c.Weight ?? DefaultWeight);

        // Calculate and set column widths, then build each cell
        foreach (var cell in Cells)
        {
            var weight = cell.Weight ?? DefaultWeight;
            var colWidth = Math.Max(MinColumnWidth, (int)(TotalWidth * weight / totalWeight));

            // Set column width for current column
            ctx.SetCurrentColumnWidth(colWidth);

            // Build the cell content
            cell.Build(ctx);

            // Move to next column
            ctx.MoveToNextColumn();
        }
    }
}
