using FluentNPOI.HotReload.Context;
using FluentNPOI.HotReload.Styling;
using FluentNPOI.HotReload.Widgets;
using FluentNPOI.Stages;
using NPOI.XSSF.UserModel;

namespace FluentNPOI.HotReload.HotReload;

/// <summary>
/// Manages the hot reload session lifecycle, including widget building,
/// file output, and refresh handling.
/// </summary>
public class HotReloadSession : IDisposable
{
    private readonly string _outputPath;
    private readonly Func<ExcelWidget> _rootWidgetFactory;
    private readonly StyleManager _styleManager = new();
    private readonly object _refreshLock = new();
    private bool _isDisposed;
    private int _refreshCount;

    /// <summary>
    /// Event raised after a successful refresh.
    /// </summary>
    public event Action<int>? RefreshCompleted;

    /// <summary>
    /// Event raised when an error occurs during refresh.
    /// </summary>
    public event Action<Exception>? RefreshError;

    /// <summary>
    /// Gets the number of refreshes performed in this session.
    /// </summary>
    public int RefreshCount => _refreshCount;

    /// <summary>
    /// Gets the output file path.
    /// </summary>
    public string OutputPath => _outputPath;

    /// <summary>
    /// Gets or sets whether to automatically exit on Rude Edit detection.
    /// When true, exits with code 42 to signal dotnet watch to restart.
    /// </summary>
    public bool AutoRestartOnRudeEdit { get; set; } = true;

    /// <summary>
    /// Gets or sets the sheet name to use for output.
    /// </summary>
    public string SheetName { get; set; } = "Sheet1";

    /// <summary>
    /// Creates a new HotReloadSession.
    /// </summary>
    /// <param name="outputPath">The path to write the Excel file to.</param>
    /// <param name="rootWidgetFactory">Factory function that creates the root widget.</param>
    public HotReloadSession(string outputPath, Func<ExcelWidget> rootWidgetFactory)
    {
        _outputPath = outputPath ?? throw new ArgumentNullException(nameof(outputPath));
        _rootWidgetFactory = rootWidgetFactory ?? throw new ArgumentNullException(nameof(rootWidgetFactory));

        // Ensure output directory exists
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    /// <summary>
    /// Starts the hot reload session, performing initial build and subscribing to hot reload events.
    /// </summary>
    public void Start()
    {
        HotReloadHandler.RefreshRequested += OnRefreshRequested;
        HotReloadHandler.RudeEditDetected += OnRudeEditDetected;

        // Initial build
        Refresh();

        Console.WriteLine($"🔥 Hot Reload session started");
        Console.WriteLine($"📊 Output: {Path.GetFullPath(_outputPath)}");
    }

    /// <summary>
    /// Stops the hot reload session and unsubscribes from events.
    /// </summary>
    public void Stop()
    {
        HotReloadHandler.RefreshRequested -= OnRefreshRequested;
        HotReloadHandler.RudeEditDetected -= OnRudeEditDetected;

        Console.WriteLine("👋 Hot Reload session stopped");
    }

    private void OnRefreshRequested(Type[]? updatedTypes)
    {
        Refresh();
    }

    private void OnRudeEditDetected()
    {
        Console.WriteLine("⚠️ Rude Edit detected - changes cannot be hot reloaded");

        if (AutoRestartOnRudeEdit)
        {
            Console.WriteLine("🔄 Triggering full restart...");
            Environment.Exit(42); // Signal dotnet watch to restart
        }
    }

    /// <summary>
    /// Performs a refresh, rebuilding the widget tree and saving the Excel file.
    /// </summary>
    public void Refresh()
    {
        lock (_refreshLock)
        {
            try
            {
                var startTime = DateTime.Now;

                // Create new workbook and reset style cache
                var workbook = new XSSFWorkbook();
                var fluentWorkbook = new FluentWorkbook(workbook);
                _styleManager.Reset(workbook);

                // Build widget tree
                var rootWidget = _rootWidgetFactory();
                var sheet = fluentWorkbook.UseSheet(SheetName, createIfMissing: true);
                var ctx = new ExcelContext(sheet, _styleManager);

                rootWidget.Build(ctx);

                // Save to file
                fluentWorkbook.SaveToPath(_outputPath);

                _refreshCount++;
                var elapsed = (DateTime.Now - startTime).TotalMilliseconds;

                Console.WriteLine($"✅ Refresh #{_refreshCount} completed in {elapsed:F0}ms");
                RefreshCompleted?.Invoke(_refreshCount);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Refresh failed: {ex.Message}");
                RefreshError?.Invoke(ex);
            }
        }
    }

    /// <summary>
    /// Disposes the session, stopping hot reload and cleaning up resources.
    /// </summary>
    public void Dispose()
    {
        if (!_isDisposed)
        {
            Stop();
            _isDisposed = true;
        }
    }
}
