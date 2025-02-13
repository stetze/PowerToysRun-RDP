using System.Diagnostics;
using System.IO;
using System.Windows.Controls;
using Microsoft.Win32;
using Microsoft.PowerToys.Settings.UI.Library;
using Wox.Plugin;
using ManagedCommon;

namespace Community.PowerToys.Run.Plugin.RDP;

public class Main : IPlugin, ISettingProvider, IReloadable, IDisposable
{
  public static string PluginID => "DF7413853DC54C2287390EE0E0C5BF42";
  public string Name => "RDP";
  public string Description => "Launches RDP connections";
  public IEnumerable<PluginAdditionalOption> AdditionalOptions =>
  [
    new()
    {
      PluginOptionType = PluginAdditionalOption.AdditionalOptionType.MultilineTextbox,
      DisplayLabel = "Predefined connections",
      DisplayDescription = "A list of connections to include in the query results by default.",
      Key = "predefinedConnections"
    }
  ];
  private bool _disposed;
  private string? _icon;
  private PluginInitContext? _context;
  private RDPConnections? _rdpConnections;
  private RDPConnections? _predefinedConnections;
  // private RDPConnectionsStore? _store;
  private SearchPhraseProvider? _searchPhraseProvider;

  /// <summary>
  /// initialize the plugin.
  /// </summary>
  /// <param name="context"></param>
  public void Init(PluginInitContext context)
  {
    _context = context;
    // _store = new RDPConnectionsStore(Path.Combine(
    //     context.CurrentPluginMetadata.PluginDirectory,
    //     "data",
    //     "connections.txt"));
    // _rdpConnections = _store.Load();
    _rdpConnections = _predefinedConnections;
    _searchPhraseProvider = new SearchPhraseProvider { Search = string.Empty };
    _context.API.ThemeChanged += OnThemeChanged;
    UpdateIconPath(_context.API.GetCurrentTheme());
  }

  private void OnThemeChanged(Theme oldTheme, Theme newTheme)
  {
    UpdateIconPath(newTheme);
  }

  private void UpdateIconPath(Theme theme)
  {
    _icon = theme is Theme.Light or Theme.HighContrastWhite ? "Images\\screen-mirroring.light.png" : "Images\\screen-mirroring.dark.png";
  }

  /// <summary>
  /// return results for the given query, starting with rpd
  /// </summary>
  /// <param name="query">search query provided by PowerToys Run</param>
  /// <returns></returns>
  public List<Result> Query(Query query)
  {
    if (_searchPhraseProvider != null)
    {
      _searchPhraseProvider.Search = query.Search;
    }
    _rdpConnections?.Reload(GetRdpConnectionsFromRegistry());

    var connections = _rdpConnections?.FindConnections(query.Search) ?? Enumerable.Empty<(string connection, int score)>();
    var predefinedConnections = _predefinedConnections?.FindConnections(query.Search) ?? Enumerable.Empty<(string connection, int score)>();
    var results = Array.Empty<Result>()
        .Union(predefinedConnections.Select(MapToResult))
        .Union(connections.Select(MapToResult))
        .ToList();

    if (results.Count == 0)
    {
      results.Add(CreateDefaultResult());
    }
    else if (query.Search.Length == 0)
    {
      results.Insert(0, CreateDefaultResult());
    }
    return results;
  }

  // private static IReadOnlyCollection<string> GetRdpConnectionsFromRegistry()
  private static string[] GetRdpConnectionsFromRegistry()
  {
    var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Terminal Server Client\Default");
    if (key is null)
    {
      return [];
    }

    return [.. key.GetValueNames().Select(x => key.GetValue(x.Trim())?.ToString()).Where(value => value != null).Cast<string>()];
  }

  private Result MapToResult((string connection, int score) item) =>
      new()
      {
        // For some reason SubTitle must be unique otherwise score is not respected
        Title = $"{item.connection}",
        SubTitle = $"Connect to {item.connection} via RDP",
        IcoPath = _icon,
        Score = item.score,
        Action = c =>
          {
            _rdpConnections?.ConnectionWasSelected(item.connection);
            // if (_rdpConnections != null)
            // {
            //   _store?.Save(_rdpConnections);
            // }

            StartMstsc(item.connection);
            return true;
          }
      };

  private Result CreateDefaultResult() =>
      new()
      {
        Title = "RDP",
        SubTitle = "Establish a new RDP connection",
        IcoPath = _icon,
        Score = 100,
        Action = c =>
          {
            if (_searchPhraseProvider != null)
            {
              StartMstsc(_searchPhraseProvider.Search);
            }
            return true;
          }
      };

  private static void StartMstsc(string connection)
  {
    if (string.IsNullOrWhiteSpace(connection))
    {
      Process.Start("mstsc");
    }
    else if (connection.Contains(".rdp") || connection.Contains(".RDP"))
    {
      Process.Start("mstsc", connection);
    }
    else
    {
      Process.Start("mstsc", "/v:" + connection);
    }
  }

  /// <summary>
  /// In order to have a newest search phrase in Result.Action lambda.
  /// Passing unwrapped Search string to Action for some reason doesn't work,
  /// since it keeps the value of first search which is empty string
  /// </summary>
  private class SearchPhraseProvider
  {
    public required string Search { get; set; }
  }

  public void ReloadData()
  {
    if (_context is null)
    {
      return;
    }
  }

  public void Dispose()
  {
    Dispose(true);
    GC.SuppressFinalize(this);
  }

  protected virtual void Dispose(bool disposing)
  {
    if (!_disposed && disposing)
    {
      _disposed = true;
    }
  }

  public Control CreateSettingPanel()
  {
    throw new NotImplementedException();
  }

  public void UpdateSettings(PowerLauncherPluginSettings settings)
  {
    _predefinedConnections = RDPConnections.Create(settings?.AdditionalOptions?.FirstOrDefault(x => x.Key == "predefinedConnections")?.TextValueAsMultilineList ?? Enumerable.Empty<string>());
  }
}