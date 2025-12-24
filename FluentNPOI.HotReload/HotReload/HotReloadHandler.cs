using System.Reflection;

[assembly: System.Reflection.Metadata.MetadataUpdateHandler(typeof(FluentNPOI.HotReload.HotReload.HotReloadHandler))]

namespace FluentNPOI.HotReload.HotReload;

/// <summary>
/// Handles .NET Hot Reload metadata updates and notifies subscribers.
/// This class is automatically invoked by the .NET runtime when code changes are applied.
/// </summary>
public static class HotReloadHandler
{
    /// <summary>
    /// Event raised when a hot reload refresh is requested.
    /// Subscribers should regenerate their Excel output.
    /// </summary>
    public static event Action<Type[]?>? RefreshRequested;

    /// <summary>
    /// Event raised when a "Rude Edit" is detected (changes that cannot be hot reloaded).
    /// Subscribers should trigger a full application restart.
    /// </summary>
    public static event Action? RudeEditDetected;

    /// <summary>
    /// Gets whether hot reload is currently active (any subscribers).
    /// </summary>
    public static bool IsActive => RefreshRequested != null;

    /// <summary>
    /// Called by the .NET runtime when hot reload successfully applies changes.
    /// </summary>
    /// <param name="updatedTypes">The types that were updated, or null if not tracked.</param>
    public static void UpdateApplication(Type[]? updatedTypes)
    {
        Console.WriteLine($"🔥 Hot Reload: {updatedTypes?.Length ?? 0} type(s) updated");
        RefreshRequested?.Invoke(updatedTypes);
    }

    /// <summary>
    /// Called by the .NET runtime when hot reload needs to clear caches.
    /// This typically indicates a "Rude Edit" that cannot be applied.
    /// </summary>
    /// <param name="updatedTypes">The types that were updated.</param>
    public static void ClearCache(Type[]? updatedTypes)
    {
        Console.WriteLine("⚠️ Hot Reload: Rude Edit detected, cache cleared");
        RudeEditDetected?.Invoke();
    }

    /// <summary>
    /// Manually triggers a refresh. Useful for testing or forced rebuilds.
    /// </summary>
    public static void TriggerRefresh()
    {
        Console.WriteLine("🔄 Manual refresh triggered");
        RefreshRequested?.Invoke(null);
    }
}
