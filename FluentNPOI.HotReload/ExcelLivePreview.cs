using FluentNPOI.HotReload.HotReload;
using FluentNPOI.HotReload.Widgets;

namespace FluentNPOI.HotReload;

/// <summary>
/// High-level API for running Excel live preview with hot reload support.
/// This is the main entry point for developers using the hot reload feature.
/// </summary>
public static class ExcelLivePreview
{
    /// <summary>
    /// Runs an Excel live preview session with the specified widget type.
    /// The session will automatically rebuild when code changes are detected.
    /// </summary>
    /// <typeparam name="TWidget">The widget type to preview (must have parameterless constructor).</typeparam>
    /// <param name="outputPath">The path to write the Excel preview file to.</param>
    /// <param name="configure">Optional configuration action for the session.</param>
    /// <example>
    /// <code>
    /// // Simple usage
    /// ExcelLivePreview.Run&lt;SalesReport&gt;("output/preview.xlsx");
    /// 
    /// // With configuration
    /// ExcelLivePreview.Run&lt;SalesReport&gt;("output/preview.xlsx", session =>
    /// {
    ///     session.SheetName = "Report";
    ///     session.AutoRestartOnRudeEdit = false;
    /// });
    /// </code>
    /// </example>
    public static void Run<TWidget>(string outputPath, Action<HotReloadSession>? configure = null)
        where TWidget : ExcelWidget, new()
    {
        Run(outputPath, () => new TWidget(), configure);
    }

    /// <summary>
    /// Runs an Excel live preview session with a custom widget factory.
    /// </summary>
    /// <param name="outputPath">The path to write the Excel preview file to.</param>
    /// <param name="widgetFactory">Factory function that creates the root widget.</param>
    /// <param name="configure">Optional configuration action for the session.</param>
    public static void Run(string outputPath, Func<ExcelWidget> widgetFactory, Action<HotReloadSession>? configure = null)
    {
        using var session = new HotReloadSession(outputPath, widgetFactory);

        // Allow custom configuration
        configure?.Invoke(session);

        // Start the session
        session.Start();

        PrintBanner(outputPath);

        // Set up graceful shutdown
        var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true;
            cts.Cancel();
            Console.WriteLine("\n🛑 Shutting down...");
        };

        // Keep the application alive
        try
        {
            Task.Delay(Timeout.Infinite, cts.Token).Wait();
        }
        catch (OperationCanceledException)
        {
            // Expected when Ctrl+C is pressed
        }
        catch (AggregateException ex) when (ex.InnerException is OperationCanceledException)
        {
            // Expected when Ctrl+C is pressed
        }

        Console.WriteLine("👋 Session ended. Goodbye!");
    }

    /// <summary>
    /// Runs an Excel live preview session asynchronously.
    /// </summary>
    /// <typeparam name="TWidget">The widget type to preview.</typeparam>
    /// <param name="outputPath">The path to write the Excel preview file to.</param>
    /// <param name="cancellationToken">Cancellation token to stop the session.</param>
    /// <param name="configure">Optional configuration action for the session.</param>
    public static async Task RunAsync<TWidget>(
        string outputPath,
        CancellationToken cancellationToken = default,
        Action<HotReloadSession>? configure = null)
        where TWidget : ExcelWidget, new()
    {
        using var session = new HotReloadSession(outputPath, () => new TWidget());

        configure?.Invoke(session);
        session.Start();

        PrintBanner(outputPath);

        try
        {
            await Task.Delay(Timeout.Infinite, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            // Expected
        }

        Console.WriteLine("👋 Session ended.");
    }

    /// <summary>
    /// Creates a HotReloadSession without starting it.
    /// Useful for advanced scenarios where you need more control.
    /// </summary>
    /// <typeparam name="TWidget">The widget type to preview.</typeparam>
    /// <param name="outputPath">The path to write the Excel preview file to.</param>
    /// <returns>A configured but not started HotReloadSession.</returns>
    public static HotReloadSession CreateSession<TWidget>(string outputPath)
        where TWidget : ExcelWidget, new()
    {
        return new HotReloadSession(outputPath, () => new TWidget());
    }

    private static void PrintBanner(string outputPath)
    {
        Console.WriteLine();
        Console.WriteLine("╔═══════════════════════════════════════════════════════════╗");
        Console.WriteLine("║           🔥 FluentNPOI Hot Reload Active 🔥              ║");
        Console.WriteLine("╠═══════════════════════════════════════════════════════════╣");
        Console.WriteLine($"║  📊 Preview: {TruncatePath(outputPath, 44),-44} ║");
        Console.WriteLine("║                                                           ║");
        Console.WriteLine("║  • Edit your widget code and save to see changes          ║");
        Console.WriteLine("║  • Press Ctrl+C to exit                                   ║");
        Console.WriteLine("╚═══════════════════════════════════════════════════════════╝");
        Console.WriteLine();
    }

    private static string TruncatePath(string path, int maxLength)
    {
        var fullPath = Path.GetFullPath(path);
        if (fullPath.Length <= maxLength)
            return fullPath;

        return "..." + fullPath.Substring(fullPath.Length - maxLength + 3);
    }
}
