using FluentNPOI.Models;
using FluentNPOI.Stages;
using FluentNPOI.HotReload.Styling;
using NPOI.SS.UserModel;

namespace FluentNPOI.HotReload.Context;

/// <summary>
/// Virtual sheet builder that accumulates operations before flushing to NPOI.
/// Tracks current position and provides fluent APIs for cell operations.
/// </summary>
public class ExcelContext
{
    private readonly FluentSheet _sheet;
    private readonly StyleManager? _styleManager;
    private int _currentRow = 1;
    private ExcelCol _currentCol = ExcelCol.A;
    private readonly int _startRow;
    private readonly ExcelCol _startCol;

    /// <summary>
    /// Gets the underlying FluentSheet being built.
    /// </summary>
    public FluentSheet Sheet => _sheet;

    /// <summary>
    /// Gets the current row position (1-indexed).
    /// </summary>
    public int CurrentRow => _currentRow;

    /// <summary>
    /// Gets the current column position.
    /// </summary>
    public ExcelCol CurrentColumn => _currentCol;

    /// <summary>
    /// Creates a new ExcelContext for building into the specified sheet.
    /// </summary>
    /// <param name="sheet">The FluentSheet to build into.</param>
    /// <param name="startRow">The starting row (1-indexed, default: 1).</param>
    /// <param name="startCol">The starting column (default: A).</param>
    public ExcelContext(FluentSheet sheet, int startRow = 1, ExcelCol startCol = ExcelCol.A)
    {
        _sheet = sheet;
        _currentRow = startRow;
        _currentCol = startCol;
        _startRow = startRow;
        _startCol = startCol;
    }

    /// <summary>
    /// Creates a new ExcelContext with StyleManager for caching styles.
    /// </summary>
    /// <param name="sheet">The FluentSheet to build into.</param>
    /// <param name="styleManager">The StyleManager for style caching.</param>
    /// <param name="startRow">The starting row (1-indexed, default: 1).</param>
    /// <param name="startCol">The starting column (default: A).</param>
    public ExcelContext(FluentSheet sheet, StyleManager styleManager, int startRow = 1, ExcelCol startCol = ExcelCol.A)
        : this(sheet, startRow, startCol)
    {
        _styleManager = styleManager;
    }

    /// <summary>
    /// Sets the value of the current cell.
    /// </summary>
    /// <param name="value">The value to set.</param>
    /// <returns>This context for chaining.</returns>
    public ExcelContext SetValue(object? value)
    {
        _sheet.SetCellPosition(_currentCol, _currentRow).SetValue(value);
        return this;
    }

    /// <summary>
    /// Sets the background color of the current cell.
    /// </summary>
    /// <param name="color">The indexed color to set.</param>
    /// <returns>This context for chaining.</returns>
    public ExcelContext SetBackgroundColor(IndexedColors color)
    {
        _sheet.SetCellPosition(_currentCol, _currentRow).SetBackgroundColor(color);
        return this;
    }

    /// <summary>
    /// Sets the font style of the current cell to bold.
    /// </summary>
    /// <returns>This context for chaining.</returns>
    public ExcelContext SetBold()
    {
        _sheet.SetCellPosition(_currentCol, _currentRow).SetFont(isBold: true);
        return this;
    }

    /// <summary>
    /// Adds a comment to the current cell.
    /// </summary>
    /// <param name="comment">The comment text.</param>
    /// <returns>This context for chaining.</returns>
    public ExcelContext SetComment(string comment)
    {
        _sheet.SetCellPosition(_currentCol, _currentRow).SetComment(comment);
        return this;
    }

    /// <summary>
    /// Applies a FluentStyle to the current cell using the StyleManager for caching.
    /// </summary>
    /// <param name="style">The FluentStyle to apply.</param>
    /// <returns>This context for chaining.</returns>
    public ExcelContext ApplyStyle(FluentStyle style)
    {
        if (!style.HasAnyStyle())
            return this;

        var cell = _sheet.SetCellPosition(_currentCol, _currentRow).GetCell();

        if (_styleManager != null)
        {
            _styleManager.ApplyStyle(cell, style);
        }
        else
        {
            // Fallback: apply style directly without caching
            ApplyStyleDirectly(cell, style);
        }

        return this;
    }

    /// <summary>
    /// Applies style directly without using StyleManager (fallback).
    /// </summary>
    private void ApplyStyleDirectly(ICell cell, FluentStyle style)
    {
        if (style.BackgroundColor != null)
        {
            _sheet.SetCellPosition(_currentCol, _currentRow).SetBackgroundColor(style.BackgroundColor);
        }

        if (style.IsBold)
        {
            _sheet.SetCellPosition(_currentCol, _currentRow).SetFont(isBold: true);
        }
    }

    /// <summary>
    /// Sets the width of the current column.
    /// </summary>
    /// <param name="col">The column to set width for.</param>
    /// <param name="width">The width in characters.</param>
    /// <returns>This context for chaining.</returns>
    public ExcelContext SetColumnWidth(ExcelCol col, int width)
    {
        _sheet.SetColumnWidth(col, width);
        return this;
    }

    /// <summary>
    /// Sets the width of the current column.
    /// </summary>
    /// <param name="width">The width in characters.</param>
    /// <returns>This context for chaining.</returns>
    public ExcelContext SetCurrentColumnWidth(int width)
    {
        _sheet.SetColumnWidth(_currentCol, width);
        return this;
    }

    /// <summary>
    /// Moves to the next row, resetting the column to the start column.
    /// </summary>
    /// <returns>This context for chaining.</returns>
    public ExcelContext MoveToNextRow()
    {
        _currentRow++;
        _currentCol = _startCol;
        return this;
    }

    /// <summary>
    /// Moves to the next column in the current row.
    /// </summary>
    /// <returns>This context for chaining.</returns>
    public ExcelContext MoveToNextColumn()
    {
        _currentCol++;
        return this;
    }

    /// <summary>
    /// Moves to a specific position.
    /// </summary>
    /// <param name="col">The column to move to.</param>
    /// <param name="row">The row to move to (1-indexed).</param>
    /// <returns>This context for chaining.</returns>
    public ExcelContext MoveTo(ExcelCol col, int row)
    {
        _currentCol = col;
        _currentRow = row;
        return this;
    }

    /// <summary>
    /// Saves the current position and creates a nested context for child widgets.
    /// </summary>
    /// <returns>A new context starting at the current position.</returns>
    public ExcelContext CreateNestedContext()
    {
        return _styleManager != null
            ? new ExcelContext(_sheet, _styleManager, _currentRow, _currentCol)
            : new ExcelContext(_sheet, _currentRow, _currentCol);
    }

    /// <summary>
    /// Resets the position to the starting point.
    /// </summary>
    /// <returns>This context for chaining.</returns>
    public ExcelContext Reset()
    {
        _currentRow = _startRow;
        _currentCol = _startCol;
        return this;
    }
}
