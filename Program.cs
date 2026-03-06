using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.IO.Ports;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Threading.Channels;
using System.Xml.Linq;
using Microsoft.AspNetCore.Http.Json;
using Microsoft.Extensions.FileProviders;

var runtimeOptions = RuntimeOptions.Parse(args);
var appDirectories = AppDirectories.Discover(Directory.GetCurrentDirectory());
var settingsStore = new SettingsStore(appDirectories.SettingsPath);
var initialSettings = settingsStore.Load();
var listenerPort = runtimeOptions.PortOverride ?? initialSettings.UiPort;

if (runtimeOptions.ValidateOnly)
{
    settingsStore.Save(initialSettings);
    Console.WriteLine("Configuration is valid.");
    Console.WriteLine($"Web UI port: {listenerPort}");
    Console.WriteLine($"Configured rules: {initialSettings.Rules.Count}");
    return;
}

var builder = WebApplication.CreateBuilder(args);
builder.WebHost.UseUrls($"http://localhost:{listenerPort}");
builder.Logging.ClearProviders();
builder.Logging.AddSimpleConsole(options =>
{
    options.SingleLine = true;
    options.TimestampFormat = "HH:mm:ss ";
});
builder.Logging.AddFilter("Microsoft", LogLevel.Warning);
builder.Logging.AddFilter("System", LogLevel.Warning);
builder.Services.Configure<JsonOptions>(options =>
{
    options.SerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
    options.SerializerOptions.WriteIndented = true;
});

builder.Services.AddSingleton(appDirectories);
builder.Services.AddSingleton(settingsStore);
builder.Services.AddSingleton(runtimeOptions);
builder.Services.AddSingleton(initialSettings);
builder.Services.AddSingleton(new ListenerInfo(listenerPort));
builder.Services.AddSingleton<TimerRouterService>();
builder.Services.AddHostedService(sp => sp.GetRequiredService<TimerRouterService>());

var app = builder.Build();
var webRootProvider = new PhysicalFileProvider(appDirectories.WebRoot);

app.MapGet("/", async context =>
{
    var file = webRootProvider.GetFileInfo("index.html");
    if (!file.Exists)
    {
        context.Response.StatusCode = StatusCodes.Status404NotFound;
        await context.Response.WriteAsync("index.html not found");
        return;
    }

    context.Response.ContentType = "text/html; charset=utf-8";
    await context.Response.SendFileAsync(file);
});

app.MapGet("/styles.css", async context =>
{
    var file = webRootProvider.GetFileInfo("styles.css");
    if (!file.Exists)
    {
        context.Response.StatusCode = StatusCodes.Status404NotFound;
        return;
    }

    context.Response.ContentType = "text/css; charset=utf-8";
    await context.Response.SendFileAsync(file);
});

app.MapGet("/app.js", async context =>
{
    var file = webRootProvider.GetFileInfo("app.js");
    if (!file.Exists)
    {
        context.Response.StatusCode = StatusCodes.Status404NotFound;
        return;
    }

    context.Response.ContentType = "application/javascript; charset=utf-8";
    await context.Response.SendFileAsync(file);
});

app.MapGet("/api/state", (TimerRouterService service) =>
{
    return Results.Json(service.GetStateSnapshot());
});

app.MapGet("/api/vmix/inputs", async (HttpContext context, TimerRouterService service) =>
{
    try
    {
        var forceRefresh = string.Equals(context.Request.Query["refresh"], "1", StringComparison.OrdinalIgnoreCase);
        var result = await service.RefreshVmixInputsAsync(forceRefresh, context.RequestAborted);
        return Results.Json(new
        {
            ok = true,
            inputs = result.Inputs,
            status = result.Status
        });
    }
    catch (Exception ex)
    {
        return JsonError(ex.Message, StatusCodes.Status500InternalServerError);
    }
});

app.MapPost("/api/settings", async (AppSettings incoming, TimerRouterService service, CancellationToken cancellationToken) =>
{
    try
    {
        var result = await service.UpdateSettingsAsync(incoming, cancellationToken);
        return Results.Json(new
        {
            ok = true,
            restartRequired = result.RestartRequired,
            state = result.State
        });
    }
    catch (SettingsValidationException ex)
    {
        return JsonError(ex.Message, StatusCodes.Status400BadRequest);
    }
    catch (Exception ex)
    {
        return JsonError(ex.Message, StatusCodes.Status500InternalServerError);
    }
});

app.MapPost("/api/rule-action", async (RuleActionRequest request, TimerRouterService service, CancellationToken cancellationToken) =>
{
    try
    {
        var state = await service.ExecuteRuleActionAsync(request, cancellationToken);
        return Results.Json(new
        {
            ok = true,
            state
        });
    }
    catch (NotFoundException ex)
    {
        return JsonError(ex.Message, StatusCodes.Status404NotFound);
    }
    catch (SettingsValidationException ex)
    {
        return JsonError(ex.Message, StatusCodes.Status400BadRequest);
    }
    catch (Exception ex)
    {
        return JsonError(ex.Message, StatusCodes.Status500InternalServerError);
    }
});

app.MapFallback(() => Results.Text("Not found", statusCode: StatusCodes.Status404NotFound));

var routerService = app.Services.GetRequiredService<TimerRouterService>();
app.Lifetime.ApplicationStarted.Register(() =>
{
    routerService.NotifyWebStarted();

    if (runtimeOptions.RunForSeconds > 0)
    {
        _ = Task.Run(async () =>
        {
            await Task.Delay(TimeSpan.FromSeconds(runtimeOptions.RunForSeconds));
            app.Lifetime.StopApplication();
        });
    }

    if (initialSettings.OpenBrowser && !runtimeOptions.NoBrowser)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = $"http://localhost:{listenerPort}/",
                UseShellExecute = true
            });
        }
        catch
        {
        }
    }
});

await app.RunAsync();

static IResult JsonError(string message, int statusCode)
{
    return Results.Json(new
    {
        ok = false,
        error = message
    }, statusCode: statusCode);
}

sealed record RuntimeOptions
{
    public bool NoBrowser { get; init; }
    public bool ValidateOnly { get; init; }
    public int RunForSeconds { get; init; }
    public int? PortOverride { get; init; }

    public static RuntimeOptions Parse(string[] args)
    {
        var options = new RuntimeOptions();
        var index = 0;

        while (index < args.Length)
        {
            var arg = args[index].Trim();
            var key = arg.TrimStart('-', '/').ToLowerInvariant();

            switch (key)
            {
                case "nobrowser":
                case "no-browser":
                    options = options with { NoBrowser = true };
                    break;
                case "validateonly":
                case "validate-only":
                    options = options with { ValidateOnly = true };
                    break;
                case "runforseconds":
                case "run-for-seconds":
                    if (index + 1 >= args.Length || !int.TryParse(args[index + 1], out var seconds))
                    {
                        throw new SettingsValidationException("run-for-seconds requires an integer value.");
                    }

                    options = options with { RunForSeconds = Math.Max(0, seconds) };
                    index += 1;
                    break;
                case "port":
                case "uiport":
                case "ui-port":
                    if (index + 1 >= args.Length || !int.TryParse(args[index + 1], out var port))
                    {
                        throw new SettingsValidationException("port requires an integer value.");
                    }

                    options = options with { PortOverride = port };
                    index += 1;
                    break;
            }

            index += 1;
        }

        return options;
    }
}

sealed class AppDirectories
{
    public required string Root { get; init; }
    public required string WebRoot { get; init; }
    public required string SettingsPath { get; init; }

    public static AppDirectories Discover(string currentDirectory)
    {
        var currentRoot = Path.GetFullPath(currentDirectory);
        var baseRoot = Path.GetFullPath(AppContext.BaseDirectory);
        var root = SelectRoot(currentRoot, baseRoot);

        return new AppDirectories
        {
            Root = root,
            WebRoot = Path.Combine(root, "web"),
            SettingsPath = Path.Combine(root, "timer.settings.json")
        };
    }

    private static string SelectRoot(params string[] candidates)
    {
        foreach (var candidate in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (File.Exists(Path.Combine(candidate, "web", "index.html")))
            {
                return candidate;
            }
        }

        return candidates.First();
    }
}

sealed class ListenerInfo
{
    public ListenerInfo(int port)
    {
        Port = port;
        Url = $"http://localhost:{port}/";
    }

    public int Port { get; }
    public string Url { get; }
}

sealed class SettingsStore
{
    private readonly string _settingsPath;
    private readonly JsonSerializerOptions _jsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public SettingsStore(string settingsPath)
    {
        _settingsPath = settingsPath;
    }

    public AppSettings Load()
    {
        if (!File.Exists(_settingsPath))
        {
            var defaults = AppSettings.CreateDefault().Normalize();
            Save(defaults);
            return defaults;
        }

        try
        {
            var text = File.ReadAllText(_settingsPath, Encoding.UTF8);
            var settings = JsonSerializer.Deserialize<AppSettings>(text, _jsonOptions) ?? AppSettings.CreateDefault();
            return settings.Normalize();
        }
        catch
        {
            var defaults = AppSettings.CreateDefault().Normalize();
            Save(defaults);
            return defaults;
        }
    }

    public void Save(AppSettings settings)
    {
        var normalized = settings.Normalize();
        var json = JsonSerializer.Serialize(normalized, _jsonOptions);
        File.WriteAllText(_settingsPath, json, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }
}

sealed class AppSettings
{
    public int UiPort { get; set; } = 8090;
    public bool OpenBrowser { get; set; } = true;
    public SerialSettings Serial { get; set; } = new();
    public VmixSettings Vmix { get; set; } = new();
    public List<TimerRule> Rules { get; set; } = new();

    public static AppSettings CreateDefault()
    {
        return new AppSettings
        {
            UiPort = 8090,
            OpenBrowser = true,
            Serial = new SerialSettings(),
            Vmix = new VmixSettings(),
            Rules = new List<TimerRule>
            {
                TimerRule.Create("Hundredths Timer", 2),
                TimerRule.Create("Thousandths Timer", 3)
            }
        };
    }

    public AppSettings Normalize()
    {
        var rules = Rules ?? new List<TimerRule>();
        return new AppSettings
        {
            UiPort = UiPort <= 0 ? 8090 : UiPort,
            OpenBrowser = OpenBrowser,
            Serial = (Serial ?? new SerialSettings()).Normalize(),
            Vmix = (Vmix ?? new VmixSettings()).Normalize(),
            Rules = rules.Count == 0
                ? new List<TimerRule> { TimerRule.Create("Hundredths Timer", 2), TimerRule.Create("Thousandths Timer", 3) }
                : rules.Select(rule => rule.Normalize()).ToList()
        };
    }
}

sealed class SerialSettings
{
    public bool Enabled { get; set; } = true;
    public string Port { get; set; } = "COM4";
    public int Baud { get; set; } = 1200;
    public int MinSendIntervalMs { get; set; } = 0;

    public SerialSettings Normalize()
    {
        return new SerialSettings
        {
            Enabled = Enabled,
            Port = (Port ?? string.Empty).Trim(),
            Baud = Baud <= 0 ? 1200 : Baud,
            MinSendIntervalMs = Math.Max(0, MinSendIntervalMs)
        };
    }
}

sealed class VmixSettings
{
    public string BaseUrl { get; set; } = "http://127.0.0.1:8088";

    public VmixSettings Normalize()
    {
        return new VmixSettings
        {
            BaseUrl = string.IsNullOrWhiteSpace(BaseUrl) ? "http://127.0.0.1:8088" : BaseUrl.Trim().TrimEnd('/')
        };
    }
}

sealed class TimerRule
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string Name { get; set; } = "Timer Rule";
    public bool Enabled { get; set; } = true;
    public int DecimalPlaces { get; set; } = -1;
    public string TargetInput { get; set; } = string.Empty;
    public string PreferredTitle { get; set; } = string.Empty;
    public string Field { get; set; } = "Time.Text";
    public int OverlayNumber { get; set; } = 1;
    public bool AutoOverlay { get; set; } = true;
    public bool UseRunningQuote { get; set; }
    public bool FlashWhenStopped { get; set; }
    public int FlashIntervalMs { get; set; } = 350;
    public int FlashDurationMs { get; set; } = 140;
    public double ZeroThreshold { get; set; } = 0.05;
    public double RearmAbove { get; set; } = 2.0;
    public double RearmBelow { get; set; } = 0.30;

    public static TimerRule Create(string name, int decimalPlaces)
    {
        return new TimerRule
        {
            Id = Guid.NewGuid().ToString(),
            Name = name,
            DecimalPlaces = decimalPlaces
        };
    }

    public TimerRule Normalize()
    {
        return new TimerRule
        {
            Id = string.IsNullOrWhiteSpace(Id) ? Guid.NewGuid().ToString() : Id.Trim(),
            Name = string.IsNullOrWhiteSpace(Name) ? "Timer Rule" : Name.Trim(),
            Enabled = Enabled,
            DecimalPlaces = DecimalPlaces,
            TargetInput = (TargetInput ?? string.Empty).Trim(),
            PreferredTitle = (PreferredTitle ?? string.Empty).Trim(),
            Field = string.IsNullOrWhiteSpace(Field) ? "Time.Text" : Field.Trim(),
            OverlayNumber = Math.Clamp(OverlayNumber, 1, 4),
            AutoOverlay = AutoOverlay,
            UseRunningQuote = UseRunningQuote,
            FlashWhenStopped = FlashWhenStopped,
            FlashIntervalMs = Math.Clamp(FlashIntervalMs, 120, 5000),
            FlashDurationMs = Math.Clamp(FlashDurationMs, 60, 5000),
            ZeroThreshold = ZeroThreshold,
            RearmAbove = RearmAbove,
            RearmBelow = RearmBelow
        };
    }
}

sealed class RuleActionRequest
{
    public string RuleId { get; set; } = string.Empty;
    public string Action { get; set; } = string.Empty;
}

sealed class StateSnapshot
{
    public ListenerInfo Listener { get; set; } = new(8090);
    public AppSettings Settings { get; set; } = AppSettings.CreateDefault();
    public LiveState Live { get; set; } = new();
    public SerialStatus SerialStatus { get; set; } = new();
    public VmixState Vmix { get; set; } = new();
    public List<RuleRuntimeSnapshot> RuleStatus { get; set; } = new();
    public List<RawSerialEntry> RawSerial { get; set; } = new();
    public List<string> Ports { get; set; } = new();
    public List<LogEntry> Logs { get; set; } = new();
}

sealed class LiveState
{
    public string LastFrame { get; set; } = string.Empty;
    public string DisplayFrame { get; set; } = string.Empty;
    public double? LastValue { get; set; }
    public int? DecimalPlaces { get; set; }
    public bool? IsRunning { get; set; }
    public string MatchedRuleId { get; set; } = string.Empty;
    public string MatchedRuleName { get; set; } = string.Empty;
    public ResolvedTarget? ResolvedTarget { get; set; }
    public string LastAction { get; set; } = "idle";
    public DateTimeOffset? UpdatedAt { get; set; }
    public string LastError { get; set; } = string.Empty;
}

sealed class SerialStatus
{
    public bool Enabled { get; set; } = true;
    public bool Connected { get; set; }
    public string Port { get; set; } = string.Empty;
    public int Baud { get; set; }
    public string Message { get; set; } = "Idle";
    public string LastError { get; set; } = string.Empty;
    public DateTimeOffset? LastConnectedAt { get; set; }
}

sealed class VmixState
{
    public bool Reachable { get; set; }
    public string BaseUrl { get; set; } = string.Empty;
    public string Message { get; set; } = "Idle";
    public string LastError { get; set; } = string.Empty;
    public DateTimeOffset? LastDiscoveryAt { get; set; }
    public List<VmixInputInfo> Inputs { get; set; } = new();
}

sealed class VmixInputInfo
{
    public string Key { get; set; } = string.Empty;
    public string Number { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string ShortTitle { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public List<VmixFieldInfo> Fields { get; set; } = new();
}

sealed class VmixFieldInfo
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
    public string Index { get; set; } = string.Empty;
}

sealed class RuleRuntimeSnapshot
{
    public string Id { get; set; } = string.Empty;
    public bool OverlayShown { get; set; }
    public string LastFrame { get; set; } = string.Empty;
    public string DisplayFrame { get; set; } = string.Empty;
    public double? LastValue { get; set; }
    public bool? IsRunning { get; set; }
    public DateTimeOffset? LastMatchedAt { get; set; }
    public DateTimeOffset? LastSentAt { get; set; }
    public DateTimeOffset? LastOverlayAt { get; set; }
    public string LastResolvedInput { get; set; } = string.Empty;
    public int Matches { get; set; }
    public ResolvedTarget? ResolvedTarget { get; set; }
}

sealed class ResolvedTarget
{
    public string Input { get; set; } = string.Empty;
    public string Source { get; set; } = string.Empty;
    public string Label { get; set; } = string.Empty;
    public string MatchedTitle { get; set; } = string.Empty;
}

sealed class LogEntry
{
    public DateTimeOffset Timestamp { get; set; } = DateTimeOffset.Now;
    public string Level { get; set; } = "info";
    public string Message { get; set; } = string.Empty;
}

sealed class RawSerialEntry
{
    public DateTimeOffset Timestamp { get; set; } = DateTimeOffset.Now;
    public string DisplayText { get; set; } = string.Empty;
    public string ParseText { get; set; } = string.Empty;
    public string Hex { get; set; } = string.Empty;
    public int Length { get; set; }
}

sealed class SerialFrame
{
    public byte[] Bytes { get; init; } = Array.Empty<byte>();
    public string DisplayText { get; init; } = string.Empty;
    public string ParseText { get; init; } = string.Empty;
    public string Hex { get; init; } = string.Empty;
}

sealed class ParsedTimerFrame
{
    public string RawText { get; init; } = string.Empty;
    public string DisplayText { get; init; } = string.Empty;
    public double? NumericValue { get; init; }
    public int? DecimalPlaces { get; init; }
    public bool IsRunning { get; init; }
}

sealed class RuleRuntimeState
{
    public bool OverlayShown { get; set; }
    public string LastFrame { get; set; } = string.Empty;
    public string DisplayFrame { get; set; } = string.Empty;
    public double? LastValue { get; set; }
    public bool? IsRunning { get; set; }
    public DateTimeOffset? LastMatchedAt { get; set; }
    public DateTimeOffset? LastSentAt { get; set; }
    public DateTimeOffset? LastOverlayAt { get; set; }
    public string LastResolvedInput { get; set; } = string.Empty;
    public int Matches { get; set; }
    public bool FlashVisible { get; set; } = true;
    public DateTimeOffset? NextFlashToggleAt { get; set; }
}

sealed class FlashUpdate
{
    public required string RuleId { get; init; }
    public required string RuleName { get; init; }
    public required string InputId { get; init; }
    public required int Alpha { get; init; }
}

sealed class VmixRefreshResult
{
    public required List<VmixInputInfo> Inputs { get; init; }
    public required VmixState Status { get; init; }
}

sealed class SettingsUpdateResult
{
    public required bool RestartRequired { get; init; }
    public required StateSnapshot State { get; init; }
}

sealed class VmixTcpResponse
{
    public string Command { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public byte[]? BinaryData { get; set; }
}

sealed class SettingsValidationException : Exception
{
    public SettingsValidationException(string message) : base(message)
    {
    }
}

sealed class NotFoundException : Exception
{
    public NotFoundException(string message) : base(message)
    {
    }
}

sealed class TimerRouterService : IHostedService, IDisposable
{
    private readonly object _gate = new();
    private readonly ILogger<TimerRouterService> _logger;
    private readonly SettingsStore _settingsStore;
    private readonly ListenerInfo _listenerInfo;
    private readonly Channel<SerialFrame> _frameChannel = Channel.CreateUnbounded<SerialFrame>(new UnboundedChannelOptions
    {
        SingleReader = true,
        SingleWriter = false
    });
    private readonly SemaphoreSlim _vmixTcpLock = new(1, 1);
    private readonly List<byte> _vmixReceiveBuffer = new();

    private CancellationTokenSource? _shutdown;
    private Task? _serialTask;
    private Task? _processorTask;
    private Task? _flashTask;
    private SerialPort? _serialPort;
    private TcpClient? _vmixTcpClient;
    private NetworkStream? _vmixStream;
    private AppSettings _settings;
    private LiveState _live = new();
    private SerialStatus _serialStatus = new();
    private VmixState _vmixState = new();
    private List<VmixInputInfo> _vmixInputs = new();
    private readonly List<LogEntry> _logs = new();
    private readonly List<RawSerialEntry> _rawSerial = new();
    private readonly Dictionary<string, RuleRuntimeState> _ruleRuntime = new();

    public TimerRouterService(
        ILogger<TimerRouterService> logger,
        SettingsStore settingsStore,
        ListenerInfo listenerInfo,
        AppSettings initialSettings)
    {
        _logger = logger;
        _settingsStore = settingsStore;
        _listenerInfo = listenerInfo;
        _settings = initialSettings.Normalize();

        _serialStatus = new SerialStatus
        {
            Enabled = _settings.Serial.Enabled,
            Connected = false,
            Port = _settings.Serial.Port,
            Baud = _settings.Serial.Baud,
            Message = "Idle",
            LastError = string.Empty
        };

        _vmixState = new VmixState
        {
            Reachable = false,
            BaseUrl = _settings.Vmix.BaseUrl,
            Message = "Idle",
            LastError = string.Empty
        };

        SyncRuleRuntime();
    }

    public Task StartAsync(CancellationToken cancellationToken)
    {
        _shutdown = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        _serialTask = Task.Run(() => SerialLoopAsync(_shutdown.Token), _shutdown.Token);
        _processorTask = Task.Run(() => FrameProcessorLoopAsync(_shutdown.Token), _shutdown.Token);
        _flashTask = Task.Run(() => FlashLoopAsync(_shutdown.Token), _shutdown.Token);
        _ = RefreshVmixInputsAsync(forceRefresh: true, cancellationToken);
        return Task.CompletedTask;
    }

    public async Task StopAsync(CancellationToken cancellationToken)
    {
        if (_shutdown is null)
        {
            return;
        }

        _shutdown.Cancel();
        CloseCurrentSerialPort();
        _frameChannel.Writer.TryComplete();

        var tasks = new[] { _serialTask, _processorTask, _flashTask }.Where(task => task is not null).Cast<Task>().ToArray();
        if (tasks.Length > 0)
        {
            try
            {
                await Task.WhenAll(tasks).WaitAsync(cancellationToken);
            }
            catch
            {
            }
        }
    }

    public void Dispose()
    {
        CloseCurrentSerialPort();
        CloseVmixConnection();
        _vmixTcpLock.Dispose();
        _shutdown?.Dispose();
    }

    public void NotifyWebStarted()
    {
        AddLog("Web UI running at " + _listenerInfo.Url);
    }

    public StateSnapshot GetStateSnapshot()
    {
        lock (_gate)
        {
            return BuildStateSnapshot();
        }
    }

    public async Task<VmixRefreshResult> RefreshVmixInputsAsync(bool forceRefresh, CancellationToken cancellationToken)
    {
        VmixState cachedStatus;
        List<VmixInputInfo> cachedInputs;

        lock (_gate)
        {
            if (!forceRefresh &&
                _vmixState.LastDiscoveryAt.HasValue &&
                _vmixInputs.Count > 0 &&
                DateTimeOffset.Now - _vmixState.LastDiscoveryAt.Value < TimeSpan.FromSeconds(15))
            {
                cachedStatus = CloneVmixState(_vmixState, _vmixInputs);
                cachedInputs = CloneInputs(_vmixInputs);
                return new VmixRefreshResult
                {
                    Inputs = cachedInputs,
                    Status = cachedStatus
                };
            }
        }

        var baseUrl = GetSettings().Vmix.BaseUrl.TrimEnd('/');
        try
        {
            var xml = await ReadVmixXmlAsync(cancellationToken);
            var inputs = ParseVmixInputs(xml);

            lock (_gate)
            {
                _vmixInputs = inputs;
                _vmixState = new VmixState
                {
                    Reachable = true,
                    BaseUrl = baseUrl,
                    Message = $"Loaded {inputs.Count} vMix inputs via TCP",
                    LastError = string.Empty,
                    LastDiscoveryAt = DateTimeOffset.Now,
                    Inputs = CloneInputs(inputs)
                };
                RebuildResolvedTargets();
                cachedStatus = CloneVmixState(_vmixState, _vmixInputs);
                cachedInputs = CloneInputs(_vmixInputs);
            }

            if (forceRefresh)
            {
                AddLog(cachedStatus.Message);
            }

            return new VmixRefreshResult
            {
                Inputs = cachedInputs,
                Status = cachedStatus
            };
        }
        catch (Exception ex)
        {
            CloseVmixConnection();

            lock (_gate)
            {
                _vmixState.Reachable = false;
                _vmixState.BaseUrl = baseUrl;
                _vmixState.Message = "vMix TCP unavailable";
                _vmixState.LastError = ex.Message;
                cachedStatus = CloneVmixState(_vmixState, _vmixInputs);
                cachedInputs = CloneInputs(_vmixInputs);
            }

            if (forceRefresh)
            {
                AddLog($"vMix refresh failed: {ex.Message}", "warn");
            }

            return new VmixRefreshResult
            {
                Inputs = cachedInputs,
                Status = cachedStatus
            };
        }
    }

    public async Task<SettingsUpdateResult> UpdateSettingsAsync(AppSettings incoming, CancellationToken cancellationToken)
    {
        if (incoming is null)
        {
            throw new SettingsValidationException("Settings payload was empty.");
        }

        var normalized = incoming.Normalize();
        var restartRequired = normalized.UiPort != _listenerInfo.Port;

        lock (_gate)
        {
            _settings = normalized;
            _serialStatus.Enabled = normalized.Serial.Enabled;
            _serialStatus.Port = normalized.Serial.Port;
            _serialStatus.Baud = normalized.Serial.Baud;
            _vmixState.BaseUrl = normalized.Vmix.BaseUrl;
            SyncRuleRuntime();
        }

        _settingsStore.Save(normalized);
        CloseCurrentSerialPort();
        CloseVmixConnection();
        await RefreshVmixInputsAsync(forceRefresh: true, cancellationToken);
        AddLog("Settings saved");

        return new SettingsUpdateResult
        {
            RestartRequired = restartRequired,
            State = GetStateSnapshot()
        };
    }

    public async Task<StateSnapshot> ExecuteRuleActionAsync(RuleActionRequest request, CancellationToken cancellationToken)
    {
        if (request is null || string.IsNullOrWhiteSpace(request.RuleId))
        {
            throw new SettingsValidationException("Rule id is required.");
        }

        TimerRule rule;
        RuleRuntimeState runtime;
        ResolvedTarget? target;

        lock (_gate)
        {
            rule = _settings.Rules.FirstOrDefault(item => string.Equals(item.Id, request.RuleId, StringComparison.Ordinal))
                ?? throw new NotFoundException("Rule not found.");
            runtime = EnsureRuleRuntime(rule.Id);
            target = ResolveRuleTarget(rule, _vmixInputs);
        }

        var action = (request.Action ?? string.Empty).Trim().ToLowerInvariant();
        switch (action)
        {
            case "overlay":
                if (target is null || string.IsNullOrWhiteSpace(target.Input))
                {
                    throw new SettingsValidationException("Rule has no vMix target.");
                }

                if (!await SendVmixFunctionAsync($"OverlayInput{rule.OverlayNumber}In", target.Input, null, null, cancellationToken))
                {
                    throw new SettingsValidationException(CurrentVmixError());
                }

                lock (_gate)
                {
                    runtime.OverlayShown = true;
                    runtime.LastOverlayAt = DateTimeOffset.Now;
                }

                AddLog($"Manual overlay trigger for {rule.Name}");
                return GetStateSnapshot();

            case "rearm":
                lock (_gate)
                {
                    runtime.OverlayShown = false;
                }

                AddLog($"Overlay rearmed manually for {rule.Name}");
                return GetStateSnapshot();

            default:
                throw new SettingsValidationException("Unsupported action.");
        }
    }

    private async Task SerialLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            var settings = GetSettings();

            if (!settings.Serial.Enabled)
            {
                UpdateSerialStatus(connected: false, "Serial disabled", string.Empty, null);
                await SafeDelay(TimeSpan.FromMilliseconds(250), cancellationToken);
                continue;
            }

            if (string.IsNullOrWhiteSpace(settings.Serial.Port))
            {
                UpdateSerialStatus(connected: false, "Choose a COM port", string.Empty, null);
                await SafeDelay(TimeSpan.FromMilliseconds(400), cancellationToken);
                continue;
            }

            SerialPort? serialPort = null;

            try
            {
                serialPort = new SerialPort(settings.Serial.Port, settings.Serial.Baud, Parity.None, 8, StopBits.One)
                {
                    NewLine = "\r",
                    ReadTimeout = 250
                };

                lock (_gate)
                {
                    _serialPort = serialPort;
                }

                serialPort.Open();
                UpdateSerialStatus(connected: true, "Connected", string.Empty, DateTimeOffset.Now);
                AddLog($"Serial connected on {settings.Serial.Port} @ {settings.Serial.Baud}");
                var lineBuffer = new List<byte>(64);

                while (!cancellationToken.IsCancellationRequested)
                {
                    int nextByte;
                    try
                    {
                        nextByte = serialPort.ReadByte();
                    }
                    catch (TimeoutException)
                    {
                        continue;
                    }

                    if (nextByte < 0)
                    {
                        continue;
                    }

                    var currentByte = (byte)nextByte;
                    if (currentByte is 0x0D or 0x0A)
                    {
                        if (lineBuffer.Count == 0)
                        {
                            continue;
                        }

                        var frame = CreateSerialFrame(lineBuffer.ToArray());
                        lineBuffer.Clear();
                        AddRawSerialEntry(frame);
                        await _frameChannel.Writer.WriteAsync(frame, cancellationToken);
                        continue;
                    }

                    lineBuffer.Add(currentByte);
                    if (lineBuffer.Count > 2048)
                    {
                        var frame = CreateSerialFrame(lineBuffer.ToArray());
                        lineBuffer.Clear();
                        AddRawSerialEntry(frame);
                        await _frameChannel.Writer.WriteAsync(frame, cancellationToken);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                UpdateSerialStatus(connected: false, "Serial reconnect pending", ex.Message, null);
                AddLog($"Serial open/read failed on {settings.Serial.Port}: {ex.Message}", "warn");
                await SafeDelay(TimeSpan.FromSeconds(1), cancellationToken);
            }
            finally
            {
                if (serialPort is not null)
                {
                    try
                    {
                        serialPort.Close();
                    }
                    catch
                    {
                    }

                    serialPort.Dispose();
                }

                lock (_gate)
                {
                    if (ReferenceEquals(_serialPort, serialPort))
                    {
                        _serialPort = null;
                    }

                    _serialStatus.Connected = false;
                }
            }
        }
    }

    private async Task FrameProcessorLoopAsync(CancellationToken cancellationToken)
    {
        try
        {
            while (await _frameChannel.Reader.WaitToReadAsync(cancellationToken))
            {
                if (!_frameChannel.Reader.TryRead(out var frame))
                {
                    continue;
                }

                // Collapse backlog so the renderer follows the newest timer value instead of old queued frames.
                while (_frameChannel.Reader.TryRead(out var newer))
                {
                    frame = newer;
                }

                await ProcessFrameAsync(frame, cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
        }
    }

    private async Task FlashLoopAsync(CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                List<FlashUpdate> pendingUpdates;
                var now = DateTimeOffset.Now;

                lock (_gate)
                {
                    pendingUpdates = new List<FlashUpdate>();

                    foreach (var rule in _settings.Rules)
                    {
                        if (!rule.Enabled || !rule.UseRunningQuote || !rule.FlashWhenStopped)
                        {
                            continue;
                        }

                        var runtime = EnsureRuleRuntime(rule.Id);
                        if (runtime.IsRunning != false || string.IsNullOrWhiteSpace(runtime.DisplayFrame))
                        {
                            continue;
                        }

                        if (!runtime.NextFlashToggleAt.HasValue || now < runtime.NextFlashToggleAt.Value)
                        {
                            continue;
                        }

                        var target = ResolveRuleTarget(rule, _vmixInputs);
                        if (target is null || string.IsNullOrWhiteSpace(target.Input))
                        {
                            continue;
                        }

                        runtime.FlashVisible = !runtime.FlashVisible;
                        var nextDelay = runtime.FlashVisible
                            ? Math.Clamp(rule.FlashIntervalMs, 120, 5000)
                            : Math.Clamp(rule.FlashDurationMs, 60, 5000);
                        runtime.NextFlashToggleAt = now.AddMilliseconds(nextDelay);

                        pendingUpdates.Add(new FlashUpdate
                        {
                            RuleId = rule.Id,
                            RuleName = rule.Name,
                            InputId = target.Input,
                            Alpha = runtime.FlashVisible ? 255 : 0
                        });
                    }
                }

                foreach (var update in pendingUpdates)
                {
                    var sent = await SendVmixFunctionAsync("SetAlpha", update.InputId, null, update.Alpha.ToString(CultureInfo.InvariantCulture), cancellationToken);
                    if (!sent)
                    {
                        continue;
                    }
                }

                await SafeDelay(TimeSpan.FromMilliseconds(40), cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
        }
    }

    private async Task ProcessFrameAsync(SerialFrame frame, CancellationToken cancellationToken)
    {
        var parsed = ParseTimerFrame(frame.ParseText);
        if (string.IsNullOrWhiteSpace(parsed.RawText))
        {
            return;
        }

        if (parsed.DecimalPlaces is null || parsed.NumericValue is null)
        {
            UpdateLiveState(parsed.RawText, parsed.DisplayText, null, null, null, null, null, "invalid", "Incoming frame was not numeric");
            return;
        }

        TimerRule? rule;
        RuleRuntimeState? runtime;
        ResolvedTarget? resolvedTarget;
        int minSendIntervalMs;
        bool shouldOverlay = false;
        bool rearmed = false;
        bool needsOpaqueAlpha = false;
        DateTimeOffset now = DateTimeOffset.Now;

        lock (_gate)
        {
            rule = MatchRule(parsed.DecimalPlaces.Value, _settings.Rules);
            if (rule is null)
            {
                UpdateLiveState(parsed.RawText, parsed.DisplayText, parsed.NumericValue, parsed.DecimalPlaces, null, null, null, "no-rule", "No enabled rule matches this decimal count");
                return;
            }

            runtime = EnsureRuleRuntime(rule.Id);
            resolvedTarget = ResolveRuleTarget(rule, _vmixInputs);
            minSendIntervalMs = _settings.Serial.MinSendIntervalMs;
            bool? runningState = rule.UseRunningQuote ? parsed.IsRunning : null;
            var previousRunningState = runtime.IsRunning;
            var previousFlashVisible = runtime.FlashVisible;
            needsOpaqueAlpha = previousFlashVisible == false &&
                (!rule.UseRunningQuote || parsed.IsRunning || !rule.FlashWhenStopped);

            if (runtime.LastFrame == parsed.RawText && !needsOpaqueAlpha)
            {
                UpdateLiveState(parsed.RawText, parsed.DisplayText, parsed.NumericValue, parsed.DecimalPlaces, runtime.IsRunning, rule, resolvedTarget, "duplicate", string.Empty);
                return;
            }

            if (resolvedTarget is null || string.IsNullOrWhiteSpace(resolvedTarget.Input))
            {
                UpdateLiveState(parsed.RawText, parsed.DisplayText, parsed.NumericValue, parsed.DecimalPlaces, runningState, rule, null, "no-target", "Rule has no vMix target");
                return;
            }

            if (rule.UseRunningQuote)
            {
                rearmed = previousRunningState == true && !parsed.IsRunning;
                if (!parsed.IsRunning)
                {
                    runtime.OverlayShown = false;
                    runtime.FlashVisible = true;
                    runtime.NextFlashToggleAt = rule.FlashWhenStopped
                        ? now.AddMilliseconds(Math.Clamp(rule.FlashIntervalMs, 120, 5000))
                        : null;
                    needsOpaqueAlpha = needsOpaqueAlpha || previousRunningState != false || previousFlashVisible == false;
                }
                else
                {
                    runtime.FlashVisible = true;
                    runtime.NextFlashToggleAt = null;
                    needsOpaqueAlpha = true;
                }
            }
            else if (runtime.LastValue is not null &&
                runtime.LastValue.Value >= rule.RearmAbove &&
                parsed.NumericValue.Value < rule.RearmBelow)
            {
                runtime.OverlayShown = false;
                rearmed = true;
            }

            if (!rule.UseRunningQuote)
            {
                runtime.FlashVisible = true;
                runtime.NextFlashToggleAt = null;
            }

            runtime.LastFrame = parsed.RawText;
            runtime.DisplayFrame = parsed.DisplayText;
            runtime.LastValue = parsed.NumericValue;
            runtime.IsRunning = runningState;
            runtime.LastMatchedAt = now;
            runtime.LastResolvedInput = resolvedTarget.Input;
            runtime.Matches += 1;

            if (runtime.LastSentAt.HasValue &&
                minSendIntervalMs > 0 &&
                now - runtime.LastSentAt.Value < TimeSpan.FromMilliseconds(minSendIntervalMs))
            {
                UpdateLiveState(parsed.RawText, parsed.DisplayText, parsed.NumericValue, parsed.DecimalPlaces, runtime.IsRunning, rule, resolvedTarget, "throttled", string.Empty);
                return;
            }

            shouldOverlay = rule.UseRunningQuote
                ? rule.AutoOverlay && parsed.IsRunning && !runtime.OverlayShown
                : rule.AutoOverlay && !runtime.OverlayShown && parsed.NumericValue.Value > rule.ZeroThreshold;
        }

        if (rearmed && rule is not null)
        {
            AddLog(rule.UseRunningQuote
                ? $"Auto-armed overlay for {rule.Name} after stop"
                : $"Rearmed overlay for {rule.Name}");
        }

        if (rule is null || runtime is null || resolvedTarget is null)
        {
            return;
        }

        var setTextSent = await SendVmixFunctionAsync("SetText", resolvedTarget.Input, rule.Field, parsed.DisplayText, cancellationToken);
        if (!setTextSent)
        {
            UpdateLiveState(parsed.RawText, parsed.DisplayText, parsed.NumericValue, parsed.DecimalPlaces, runtime.IsRunning, rule, resolvedTarget, "vmix-error", CurrentVmixError());
            return;
        }

        lock (_gate)
        {
            runtime.LastSentAt = now;
        }

        if (needsOpaqueAlpha)
        {
            var alphaSent = await SendVmixFunctionAsync("SetAlpha", resolvedTarget.Input, null, "255", cancellationToken);
            if (!alphaSent)
            {
                UpdateLiveState(parsed.RawText, parsed.DisplayText, parsed.NumericValue, parsed.DecimalPlaces, runtime.IsRunning, rule, resolvedTarget, "vmix-error", CurrentVmixError());
                return;
            }
        }

        if (shouldOverlay)
        {
            var overlaySent = await SendVmixFunctionAsync($"OverlayInput{rule.OverlayNumber}In", resolvedTarget.Input, null, null, cancellationToken);
            if (overlaySent)
            {
                lock (_gate)
                {
                    runtime.OverlayShown = true;
                    runtime.LastOverlayAt = DateTimeOffset.Now;
                }

                AddLog($"Overlay {rule.OverlayNumber} triggered for {rule.Name}");
            }
        }

        UpdateLiveState(parsed.RawText, parsed.DisplayText, parsed.NumericValue, parsed.DecimalPlaces, runtime.IsRunning, rule, resolvedTarget, "sent", string.Empty);
    }

    private async Task<bool> SendVmixFunctionAsync(string functionName, string inputId, string? fieldName, string? value, CancellationToken cancellationToken)
    {
        var settings = GetSettings();
        var queryParts = new List<string>
        {
            "Input=" + Uri.EscapeDataString(inputId)
        };

        if (!string.IsNullOrWhiteSpace(fieldName))
        {
            queryParts.Add("SelectedName=" + Uri.EscapeDataString(fieldName));
        }

        if (!string.IsNullOrWhiteSpace(value))
        {
            queryParts.Add("Value=" + Uri.EscapeDataString(value));
        }

        try
        {
            var response = await SendTcpCommandAsync(
                "FUNCTION " + functionName + " " + string.Join("&", queryParts),
                expectsBinaryData: false,
                cancellationToken);

            lock (_gate)
            {
                _vmixState.Reachable = true;
                _vmixState.BaseUrl = settings.Vmix.BaseUrl;
                _vmixState.Message = "vMix TCP connected";
                _vmixState.LastError = string.Empty;
            }

            if (!response.Command.Equals("FUNCTION", StringComparison.OrdinalIgnoreCase) ||
                !response.Status.Equals("OK", StringComparison.OrdinalIgnoreCase))
            {
                throw new IOException(string.IsNullOrWhiteSpace(response.Message) ? "vMix rejected the TCP command." : response.Message);
            }

            return true;
        }
        catch (Exception ex)
        {
            CloseVmixConnection();

            lock (_gate)
            {
                _vmixState.Reachable = false;
                _vmixState.BaseUrl = settings.Vmix.BaseUrl;
                _vmixState.Message = "vMix TCP send failed";
                _vmixState.LastError = ex.Message;
            }

            AddLog($"vMix TCP send failed: {ex.Message}", "warn");
            return false;
        }
    }

    private async Task<string> ReadVmixXmlAsync(CancellationToken cancellationToken)
    {
        var response = await SendTcpCommandAsync("XML", expectsBinaryData: true, cancellationToken);
        if (!response.Command.Equals("XML", StringComparison.OrdinalIgnoreCase))
        {
            throw new IOException("Unexpected TCP response while requesting XML.");
        }

        if (response.BinaryData is null)
        {
            throw new IOException("vMix TCP XML response did not contain payload data.");
        }

        return Encoding.UTF8.GetString(response.BinaryData).Trim('\0', '\r', '\n');
    }

    private async Task<VmixTcpResponse> SendTcpCommandAsync(string command, bool expectsBinaryData, CancellationToken cancellationToken)
    {
        await _vmixTcpLock.WaitAsync(cancellationToken);
        try
        {
            await EnsureVmixConnectionAsync(cancellationToken);
            await WriteTcpLineAsync(command, cancellationToken);

            while (true)
            {
                var line = await ReadTcpLineAsync(cancellationToken);
                if (line.StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var tokens = line.Split(' ', 3, StringSplitOptions.None);
                if (tokens.Length < 2)
                {
                    throw new IOException("Malformed response from vMix TCP API.");
                }

                if (expectsBinaryData && int.TryParse(tokens[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var binaryLength))
                {
                    var payload = await ReadTcpBytesAsync(binaryLength, cancellationToken);
                    return new VmixTcpResponse
                    {
                        Command = tokens[0],
                        Status = tokens[1],
                        Message = tokens.Length >= 3 ? tokens[2] : string.Empty,
                        BinaryData = payload
                    };
                }

                return new VmixTcpResponse
                {
                    Command = tokens[0],
                    Status = tokens[1],
                    Message = tokens.Length >= 3 ? tokens[2] : string.Empty
                };
            }
        }
        finally
        {
            _vmixTcpLock.Release();
        }
    }

    private async Task EnsureVmixConnectionAsync(CancellationToken cancellationToken)
    {
        if (_vmixTcpClient is { Connected: true } && _vmixStream is not null)
        {
            return;
        }

        CloseVmixConnection();

        var host = GetVmixTcpHost(GetSettings().Vmix.BaseUrl);
        var client = new TcpClient
        {
            NoDelay = true
        };

        using var connectCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        connectCts.CancelAfter(TimeSpan.FromMilliseconds(400));
        await client.ConnectAsync(host, 8099, connectCts.Token);
        client.NoDelay = true;

        _vmixTcpClient = client;
        _vmixStream = client.GetStream();
        _vmixReceiveBuffer.Clear();

        _ = await TryReadTcpLineAsync(TimeSpan.FromMilliseconds(100), cancellationToken);
    }

    private void CloseVmixConnection()
    {
        try
        {
            _vmixStream?.Dispose();
        }
        catch
        {
        }

        try
        {
            _vmixTcpClient?.Close();
        }
        catch
        {
        }

        _vmixStream = null;
        _vmixTcpClient = null;
        _vmixReceiveBuffer.Clear();
    }

    private async Task WriteTcpLineAsync(string line, CancellationToken cancellationToken)
    {
        if (_vmixStream is null)
        {
            throw new IOException("vMix TCP stream is not connected.");
        }

        var payload = Encoding.UTF8.GetBytes(line + "\r\n");
        await _vmixStream.WriteAsync(payload, cancellationToken);
        await _vmixStream.FlushAsync(cancellationToken);
    }

    private async Task<string?> TryReadTcpLineAsync(TimeSpan timeout, CancellationToken cancellationToken)
    {
        using var readCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        readCts.CancelAfter(timeout);
        try
        {
            return await ReadTcpLineAsync(readCts.Token);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            return null;
        }
    }

    private async Task<string> ReadTcpLineAsync(CancellationToken cancellationToken)
    {
        while (true)
        {
            var existing = TryExtractBufferedLine();
            if (existing is not null)
            {
                return existing;
            }

            await FillTcpBufferAsync(minimumBytes: 1, cancellationToken);
        }
    }

    private async Task<byte[]> ReadTcpBytesAsync(int count, CancellationToken cancellationToken)
    {
        while (_vmixReceiveBuffer.Count < count)
        {
            await FillTcpBufferAsync(count, cancellationToken);
        }

        var payload = _vmixReceiveBuffer.Take(count).ToArray();
        _vmixReceiveBuffer.RemoveRange(0, count);
        return payload;
    }

    private async Task FillTcpBufferAsync(int minimumBytes, CancellationToken cancellationToken)
    {
        if (_vmixStream is null)
        {
            throw new IOException("vMix TCP stream is not connected.");
        }

        var chunk = new byte[4096];
        while (_vmixReceiveBuffer.Count < minimumBytes)
        {
            using var readCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            readCts.CancelAfter(TimeSpan.FromMilliseconds(500));
            var bytesRead = await _vmixStream.ReadAsync(chunk.AsMemory(0, chunk.Length), readCts.Token);
            if (bytesRead <= 0)
            {
                throw new IOException("vMix TCP connection closed.");
            }

            _vmixReceiveBuffer.AddRange(chunk.AsSpan(0, bytesRead).ToArray());
            if (bytesRead < chunk.Length)
            {
                break;
            }
        }
    }

    private string? TryExtractBufferedLine()
    {
        for (var index = 0; index < _vmixReceiveBuffer.Count - 1; index += 1)
        {
            if (_vmixReceiveBuffer[index] == '\r' && _vmixReceiveBuffer[index + 1] == '\n')
            {
                var lineBytes = _vmixReceiveBuffer.Take(index).ToArray();
                _vmixReceiveBuffer.RemoveRange(0, index + 2);
                return Encoding.UTF8.GetString(lineBytes);
            }
        }

        return null;
    }

    private static string GetVmixTcpHost(string baseUrl)
    {
        if (Uri.TryCreate(baseUrl, UriKind.Absolute, out var uri))
        {
            return uri.Host;
        }

        var trimmed = baseUrl.Trim();
        if (trimmed.Contains("://", StringComparison.Ordinal))
        {
            trimmed = trimmed[(trimmed.IndexOf("://", StringComparison.Ordinal) + 3)..];
        }

        var slashIndex = trimmed.IndexOf('/');
        if (slashIndex >= 0)
        {
            trimmed = trimmed[..slashIndex];
        }

        var colonIndex = trimmed.IndexOf(':');
        if (colonIndex >= 0)
        {
            trimmed = trimmed[..colonIndex];
        }

        return string.IsNullOrWhiteSpace(trimmed) ? "127.0.0.1" : trimmed;
    }

    private void CloseCurrentSerialPort()
    {
        SerialPort? serialPort;
        lock (_gate)
        {
            serialPort = _serialPort;
            _serialPort = null;
        }

        if (serialPort is null)
        {
            return;
        }

        try
        {
            serialPort.Close();
        }
        catch
        {
        }

        serialPort.Dispose();
    }

    private void UpdateSerialStatus(bool connected, string message, string lastError, DateTimeOffset? connectedAt)
    {
        lock (_gate)
        {
            _serialStatus.Enabled = _settings.Serial.Enabled;
            _serialStatus.Connected = connected;
            _serialStatus.Port = _settings.Serial.Port;
            _serialStatus.Baud = _settings.Serial.Baud;
            _serialStatus.Message = message;
            _serialStatus.LastError = lastError;
            _serialStatus.LastConnectedAt = connectedAt;
        }
    }

    private void UpdateLiveState(string frame, string displayFrame, double? value, int? decimalPlaces, bool? isRunning, TimerRule? rule, ResolvedTarget? resolvedTarget, string action, string error)
    {
        lock (_gate)
        {
            _live = new LiveState
            {
                LastFrame = frame,
                DisplayFrame = displayFrame,
                LastValue = value,
                DecimalPlaces = decimalPlaces,
                IsRunning = isRunning,
                MatchedRuleId = rule?.Id ?? string.Empty,
                MatchedRuleName = rule?.Name ?? string.Empty,
                ResolvedTarget = resolvedTarget is null ? null : CloneResolvedTarget(resolvedTarget),
                LastAction = action,
                UpdatedAt = DateTimeOffset.Now,
                LastError = error
            };
        }
    }

    private AppSettings GetSettings()
    {
        lock (_gate)
        {
            return _settings.Normalize();
        }
    }

    private void SyncRuleRuntime()
    {
        var validIds = new HashSet<string>(_settings.Rules.Select(rule => rule.Id), StringComparer.Ordinal);
        foreach (var id in validIds)
        {
            EnsureRuleRuntime(id);
        }

        var staleIds = _ruleRuntime.Keys.Where(id => !validIds.Contains(id)).ToList();
        foreach (var id in staleIds)
        {
            _ruleRuntime.Remove(id);
        }
    }

    private RuleRuntimeState EnsureRuleRuntime(string ruleId)
    {
        if (!_ruleRuntime.TryGetValue(ruleId, out var runtime))
        {
            runtime = new RuleRuntimeState();
            _ruleRuntime[ruleId] = runtime;
        }

        return runtime;
    }

    private void AddLog(string message, string level = "info")
    {
        lock (_gate)
        {
            _logs.Insert(0, new LogEntry
            {
                Timestamp = DateTimeOffset.Now,
                Level = level,
                Message = message
            });

            while (_logs.Count > 150)
            {
                _logs.RemoveAt(_logs.Count - 1);
            }
        }

        _logger.LogInformation("[{Level}] {Message}", level.ToUpperInvariant(), message);
    }

    private static async Task SafeDelay(TimeSpan delay, CancellationToken cancellationToken)
    {
        try
        {
            await Task.Delay(delay, cancellationToken);
        }
        catch (OperationCanceledException)
        {
        }
    }

    private static SerialFrame CreateSerialFrame(byte[] bytes)
    {
        return new SerialFrame
        {
            Bytes = bytes,
            DisplayText = BuildSerialDisplayText(bytes),
            ParseText = Encoding.ASCII.GetString(bytes),
            Hex = string.Join(" ", bytes.Select(value => value.ToString("X2", CultureInfo.InvariantCulture)))
        };
    }

    private static string BuildSerialDisplayText(byte[] bytes)
    {
        var builder = new StringBuilder(bytes.Length * 2);
        foreach (var value in bytes)
        {
            if (value is >= 32 and <= 126)
            {
                builder.Append((char)value);
                continue;
            }

            if (value == 9)
            {
                builder.Append("\\t");
                continue;
            }

            builder.Append("\\x");
            builder.Append(value.ToString("X2", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private void AddRawSerialEntry(SerialFrame frame)
    {
        lock (_gate)
        {
            _rawSerial.Insert(0, new RawSerialEntry
            {
                Timestamp = DateTimeOffset.Now,
                DisplayText = frame.DisplayText,
                ParseText = frame.ParseText,
                Hex = frame.Hex,
                Length = frame.Bytes.Length
            });

            while (_rawSerial.Count > 120)
            {
                _rawSerial.RemoveAt(_rawSerial.Count - 1);
            }
        }
    }

    private StateSnapshot BuildStateSnapshot()
    {
        var snapshot = new StateSnapshot
        {
            Listener = _listenerInfo,
            Settings = _settings.Normalize(),
            Live = CloneLiveState(_live),
            SerialStatus = new SerialStatus
            {
                Enabled = _serialStatus.Enabled,
                Connected = _serialStatus.Connected,
                Port = _serialStatus.Port,
                Baud = _serialStatus.Baud,
                Message = _serialStatus.Message,
                LastError = _serialStatus.LastError,
                LastConnectedAt = _serialStatus.LastConnectedAt
            },
            Vmix = CloneVmixState(_vmixState, _vmixInputs),
            RawSerial = _rawSerial.Select(entry => new RawSerialEntry
            {
                Timestamp = entry.Timestamp,
                DisplayText = entry.DisplayText,
                ParseText = entry.ParseText,
                Hex = entry.Hex,
                Length = entry.Length
            }).ToList(),
            Ports = SerialPort.GetPortNames().OrderBy(port => port, StringComparer.OrdinalIgnoreCase).ToList(),
            Logs = _logs.Select(entry => new LogEntry
            {
                Timestamp = entry.Timestamp,
                Level = entry.Level,
                Message = entry.Message
            }).ToList()
        };

        snapshot.RuleStatus = _settings.Rules.Select(rule =>
        {
            var runtime = EnsureRuleRuntime(rule.Id);
            return new RuleRuntimeSnapshot
            {
                Id = rule.Id,
                OverlayShown = runtime.OverlayShown,
                LastFrame = runtime.LastFrame,
                DisplayFrame = runtime.DisplayFrame,
                LastValue = runtime.LastValue,
                IsRunning = runtime.IsRunning,
                LastMatchedAt = runtime.LastMatchedAt,
                LastSentAt = runtime.LastSentAt,
                LastOverlayAt = runtime.LastOverlayAt,
                LastResolvedInput = runtime.LastResolvedInput,
                Matches = runtime.Matches,
                ResolvedTarget = ResolveRuleTarget(rule, _vmixInputs)
            };
        }).ToList();

        return snapshot;
    }

    private static LiveState CloneLiveState(LiveState source)
    {
        return new LiveState
        {
            LastFrame = source.LastFrame,
            DisplayFrame = source.DisplayFrame,
            LastValue = source.LastValue,
            DecimalPlaces = source.DecimalPlaces,
            IsRunning = source.IsRunning,
            MatchedRuleId = source.MatchedRuleId,
            MatchedRuleName = source.MatchedRuleName,
            ResolvedTarget = source.ResolvedTarget is null ? null : CloneResolvedTarget(source.ResolvedTarget),
            LastAction = source.LastAction,
            UpdatedAt = source.UpdatedAt,
            LastError = source.LastError
        };
    }

    private static VmixState CloneVmixState(VmixState source, List<VmixInputInfo> inputs)
    {
        return new VmixState
        {
            Reachable = source.Reachable,
            BaseUrl = source.BaseUrl,
            Message = source.Message,
            LastError = source.LastError,
            LastDiscoveryAt = source.LastDiscoveryAt,
            Inputs = CloneInputs(inputs)
        };
    }

    private static List<VmixInputInfo> CloneInputs(List<VmixInputInfo> inputs)
    {
        return inputs.Select(input => new VmixInputInfo
        {
            Key = input.Key,
            Number = input.Number,
            Title = input.Title,
            ShortTitle = input.ShortTitle,
            Type = input.Type,
            Fields = input.Fields.Select(field => new VmixFieldInfo
            {
                Name = field.Name,
                Value = field.Value,
                Index = field.Index
            }).ToList()
        }).ToList();
    }

    private static ResolvedTarget CloneResolvedTarget(ResolvedTarget source)
    {
        return new ResolvedTarget
        {
            Input = source.Input,
            Source = source.Source,
            Label = source.Label,
            MatchedTitle = source.MatchedTitle
        };
    }

    private void RebuildResolvedTargets()
    {
        foreach (var rule in _settings.Rules)
        {
            var runtime = EnsureRuleRuntime(rule.Id);
            var target = ResolveRuleTarget(rule, _vmixInputs);
            runtime.LastResolvedInput = target?.Input ?? runtime.LastResolvedInput;
        }
    }

    private static TimerRule? MatchRule(int decimalPlaces, List<TimerRule> rules)
    {
        foreach (var rule in rules)
        {
            if (rule.Enabled && rule.DecimalPlaces == decimalPlaces)
            {
                return rule;
            }
        }

        foreach (var rule in rules)
        {
            if (rule.Enabled && rule.DecimalPlaces < 0)
            {
                return rule;
            }
        }

        return null;
    }

    private static ParsedTimerFrame ParseTimerFrame(string text)
    {
        var rawText = (text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(rawText))
        {
            return new ParsedTimerFrame();
        }

        var isRunning = rawText.EndsWith("\"", StringComparison.Ordinal);
        var displayText = isRunning
            ? rawText[..^1].TrimEnd()
            : rawText;

        return new ParsedTimerFrame
        {
            RawText = rawText,
            DisplayText = displayText,
            NumericValue = TryParseNumeric(displayText),
            DecimalPlaces = GetDecimalPlaces(displayText),
            IsRunning = isRunning
        };
    }

    private static int? GetDecimalPlaces(string text)
    {
        var dotIndex = text.IndexOfAny(['.', ',']);
        if (dotIndex < 0)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out _) ? 0 : null;
        }

        return text.Length - dotIndex - 1;
    }

    private static double? TryParseNumeric(string text)
    {
        var normalized = text.Trim().Replace(',', '.');
        return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;
    }

    private static ResolvedTarget? ResolveRuleTarget(TimerRule rule, List<VmixInputInfo> inputs)
    {
        if (!string.IsNullOrWhiteSpace(rule.TargetInput))
        {
            var known = inputs.FirstOrDefault(input =>
                string.Equals(input.Key, rule.TargetInput, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(input.Number, rule.TargetInput, StringComparison.OrdinalIgnoreCase));

            return new ResolvedTarget
            {
                Input = rule.TargetInput,
                Source = "explicit",
                Label = known is null ? rule.TargetInput : GetInputLabel(known),
                MatchedTitle = known?.Title ?? string.Empty
            };
        }

        var term = string.IsNullOrWhiteSpace(rule.PreferredTitle) ? rule.Name : rule.PreferredTitle;
        if (string.IsNullOrWhiteSpace(term))
        {
            return null;
        }

        var match = FindBestInputMatch(term.Trim(), inputs);
        if (match is null)
        {
            return null;
        }

        return new ResolvedTarget
        {
            Input = string.IsNullOrWhiteSpace(match.Key) ? match.Number : match.Key,
            Source = "smart",
            Label = GetInputLabel(match),
            MatchedTitle = match.Title
        };
    }

    private static VmixInputInfo? FindBestInputMatch(string term, List<VmixInputInfo> inputs)
    {
        var lowered = term.ToLowerInvariant();
        VmixInputInfo? best = null;
        var bestScore = -1;

        foreach (var input in inputs)
        {
            var candidates = new[]
            {
                input.Number,
                input.Title,
                input.ShortTitle
            };

            var score = -1;
            foreach (var candidate in candidates)
            {
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    continue;
                }

                var text = candidate.Trim().ToLowerInvariant();
                if (text == lowered)
                {
                    score = Math.Max(score, 100);
                }
                else if (text.StartsWith(lowered, StringComparison.Ordinal))
                {
                    score = Math.Max(score, 85);
                }
                else if (text.Contains(lowered, StringComparison.Ordinal))
                {
                    score = Math.Max(score, 70);
                }
            }

            if (score > bestScore)
            {
                bestScore = score;
                best = input;
            }
        }

        return best;
    }

    private static string GetInputLabel(VmixInputInfo input)
    {
        if (!string.IsNullOrWhiteSpace(input.Title) && !string.IsNullOrWhiteSpace(input.ShortTitle) && !string.Equals(input.Title, input.ShortTitle, StringComparison.Ordinal))
        {
            return $"#{input.Number} {input.Title} ({input.ShortTitle})";
        }

        if (!string.IsNullOrWhiteSpace(input.Title))
        {
            return $"#{input.Number} {input.Title}";
        }

        if (!string.IsNullOrWhiteSpace(input.ShortTitle))
        {
            return $"#{input.Number} {input.ShortTitle}";
        }

        return $"#{input.Number}";
    }

    private static List<VmixInputInfo> ParseVmixInputs(string xml)
    {
        var document = XDocument.Parse(xml);
        var inputs = document.Descendants("input")
            .Select(node => new VmixInputInfo
            {
                Key = (string?)node.Attribute("key") ?? string.Empty,
                Number = (string?)node.Attribute("number") ?? string.Empty,
                Title = (string?)node.Attribute("title") ?? string.Empty,
                ShortTitle = (string?)node.Attribute("shortTitle") ?? string.Empty,
                Type = (string?)node.Attribute("type") ?? string.Empty,
                Fields = node.Elements("text").Select(text => new VmixFieldInfo
                {
                    Name = (string?)text.Attribute("name") ?? string.Empty,
                    Value = text.Value ?? string.Empty,
                    Index = (string?)text.Attribute("index") ?? string.Empty
                }).ToList()
            })
            .OrderBy(input => int.TryParse(input.Number, out var number) ? number : int.MaxValue)
            .ThenBy(input => input.Number, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return inputs;
    }

    private string CurrentVmixError()
    {
        lock (_gate)
        {
            return string.IsNullOrWhiteSpace(_vmixState.LastError) ? "vMix TCP request failed." : _vmixState.LastError;
        }
    }
}
