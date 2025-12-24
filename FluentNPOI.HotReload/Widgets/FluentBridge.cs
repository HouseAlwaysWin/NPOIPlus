using System.Runtime.CompilerServices;
using FluentNPOI.HotReload.Context;
using FluentNPOI.Stages;

namespace FluentNPOI.HotReload.Widgets;

/// <summary>
/// A bridge widget that allows using the original FluentNPOI imperative API
/// within the declarative widget system. This enables gradual migration and
/// hybrid development styles.
/// </summary>
/// <example>
/// <code>
/// // Use existing FluentNPOI code within widgets
/// new Column(
///     new Header("Report Title"),
///     new FluentBridge(sheet =>
///     {
///         // Original FluentNPOI style
///         sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Hello");
///         sheet.SetCellPosition(ExcelCol.B, 1).SetValue("World");
///     }),
///     new Label("Footer")
/// )
/// </code>
/// </example>
public class FluentBridge : ExcelWidget
{
    private readonly Action<FluentSheet> _buildAction;
    private readonly Action<FluentSheet, int, int>? _buildActionWithPosition;

    /// <summary>
    /// Creates a FluentBridge widget with simple sheet access.
    /// </summary>
    /// <param name="buildAction">Action to execute with the FluentSheet.</param>
    /// <param name="filePath">Auto-captured source file path.</param>
    /// <param name="lineNumber">Auto-captured line number.</param>
    public FluentBridge(
        Action<FluentSheet> buildAction,
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0)
        : base(filePath, lineNumber)
    {
        _buildAction = buildAction ?? throw new ArgumentNullException(nameof(buildAction));
    }

    /// <summary>
    /// Creates a FluentBridge widget with position-aware access.
    /// Receives the current row and column as parameters.
    /// </summary>
    /// <param name="buildAction">Action to execute with sheet and current position.</param>
    /// <param name="filePath">Auto-captured source file path.</param>
    /// <param name="lineNumber">Auto-captured line number.</param>
    public FluentBridge(
        Action<FluentSheet, int, int> buildAction,
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0)
        : base(filePath, lineNumber)
    {
        _buildActionWithPosition = buildAction ?? throw new ArgumentNullException(nameof(buildAction));
        _buildAction = _ => { }; // Placeholder
    }

    /// <inheritdoc/>
    public override void Build(ExcelContext ctx)
    {
        if (_buildActionWithPosition != null)
        {
            _buildActionWithPosition(ctx.Sheet, ctx.CurrentRow, (int)ctx.CurrentColumn);
        }
        else
        {
            _buildAction(ctx.Sheet);
        }
    }
}

/// <summary>
/// A widget that wraps any FluentNPOI workbook-level operation.
/// Useful for operations that need access to the full workbook.
/// </summary>
public class WorkbookBridge : ExcelWidget
{
    private readonly Action<FluentWorkbook> _buildAction;
    private FluentWorkbook? _workbook;

    /// <summary>
    /// Creates a WorkbookBridge widget.
    /// </summary>
    /// <param name="buildAction">Action to execute with the FluentWorkbook.</param>
    /// <param name="filePath">Auto-captured source file path.</param>
    /// <param name="lineNumber">Auto-captured line number.</param>
    public WorkbookBridge(
        Action<FluentWorkbook> buildAction,
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0)
        : base(filePath, lineNumber)
    {
        _buildAction = buildAction ?? throw new ArgumentNullException(nameof(buildAction));
    }

    /// <summary>
    /// Sets the workbook reference. Must be called before Build.
    /// </summary>
    public WorkbookBridge SetWorkbook(FluentWorkbook workbook)
    {
        _workbook = workbook;
        return this;
    }

    /// <inheritdoc/>
    public override void Build(ExcelContext ctx)
    {
        if (_workbook != null)
        {
            _buildAction(_workbook);
        }
        else
        {
            throw new InvalidOperationException("Workbook not set. Call SetWorkbook() before Build().");
        }
    }
}

/// <summary>
/// Extension methods for integrating FluentNPOI with the widget system.
/// </summary>
public static class FluentIntegrationExtensions
{
    /// <summary>
    /// Creates a widget from a FluentSheet action.
    /// </summary>
    /// <param name="action">The action to execute.</param>
    /// <returns>A FluentBridge widget.</returns>
    public static FluentBridge ToWidget(this Action<FluentSheet> action)
    {
        return new FluentBridge(action);
    }

    /// <summary>
    /// Creates a widget from a position-aware FluentSheet action.
    /// </summary>
    /// <param name="action">The action to execute with row and column.</param>
    /// <returns>A FluentBridge widget.</returns>
    public static FluentBridge ToWidget(this Action<FluentSheet, int, int> action)
    {
        return new FluentBridge(action);
    }

    /// <summary>
    /// Builds a widget tree into an existing FluentSheet at the specified position.
    /// </summary>
    /// <param name="sheet">The FluentSheet to build into.</param>
    /// <param name="widget">The widget tree to build.</param>
    /// <param name="startRow">The starting row (1-indexed).</param>
    /// <param name="startCol">The starting column.</param>
    public static FluentSheet BuildWidget(
        this FluentSheet sheet,
        ExcelWidget widget,
        int startRow = 1,
        FluentNPOI.Models.ExcelCol startCol = FluentNPOI.Models.ExcelCol.A)
    {
        var ctx = new ExcelContext(sheet, startRow, startCol);
        widget.Build(ctx);
        return sheet;
    }
}
