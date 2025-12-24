using System.Runtime.CompilerServices;
using FluentNPOI.HotReload.Context;
using FluentNPOI.HotReload.Diagnostics;
using FluentNPOI.HotReload.Styling;

namespace FluentNPOI.HotReload.Widgets;

/// <summary>
/// Base class for all declarative Excel widgets.
/// Widgets follow a Flutter-style composition pattern for building Excel content.
/// </summary>
public abstract class ExcelWidget
{
    /// <summary>
    /// Optional key for identifying this widget during state management and hot reload.
    /// Widgets with keys can preserve user-entered values across rebuilds.
    /// </summary>
    public string? Key { get; set; }

    /// <summary>
    /// The source location where this widget was created.
    /// Used for debugging and error reporting.
    /// </summary>
    public DebugLocation SourceLocation { get; }

    /// <summary>
    /// Weight for Flexbox-style layout calculation.
    /// Used by FlexibleRow to automatically calculate column widths.
    /// </summary>
    public int? Weight { get; set; }

    /// <summary>
    /// The style to apply to this widget.
    /// </summary>
    public FluentStyle? Style { get; set; }

    /// <summary>
    /// Creates a new ExcelWidget, automatically capturing the source location.
    /// </summary>
    /// <param name="filePath">The source file path (auto-captured).</param>
    /// <param name="lineNumber">The line number (auto-captured).</param>
    protected ExcelWidget(
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0)
    {
        SourceLocation = new DebugLocation(filePath, lineNumber);
    }

    /// <summary>
    /// Builds the widget content into the Excel context.
    /// </summary>
    /// <param name="ctx">The Excel context to build into.</param>
    public abstract void Build(ExcelContext ctx);

    /// <summary>
    /// Sets the key for this widget and returns itself for fluent chaining.
    /// </summary>
    /// <param name="key">The key to identify this widget.</param>
    /// <returns>This widget instance.</returns>
    public ExcelWidget WithKey(string key)
    {
        Key = key;
        return this;
    }

    /// <summary>
    /// Sets the weight for Flexbox-style layout and returns itself for fluent chaining.
    /// </summary>
    /// <param name="weight">The weight value for layout calculation.</param>
    /// <returns>This widget instance.</returns>
    public ExcelWidget WithWeight(int weight)
    {
        Weight = weight;
        return this;
    }

    /// <summary>
    /// Sets the style for this widget and returns itself for fluent chaining.
    /// </summary>
    /// <param name="style">The FluentStyle to apply.</param>
    /// <returns>This widget instance.</returns>
    public ExcelWidget WithStyle(FluentStyle style)
    {
        Style = style;
        return this;
    }

    /// <summary>
    /// Applies the widget's style to the current cell if a style is set.
    /// </summary>
    /// <param name="ctx">The Excel context.</param>
    protected void ApplyStyleIfSet(ExcelContext ctx)
    {
        if (Style != null)
        {
            ctx.ApplyStyle(Style);
        }
    }
}
