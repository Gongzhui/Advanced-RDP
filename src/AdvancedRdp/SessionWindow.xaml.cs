using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Threading;
using AdvancedRdp.Controls;
using AdvancedRdp.Models;
using AdvancedRdp.Services;

namespace AdvancedRdp;

public partial class SessionWindow : Window
{
    private readonly HostEntry _entry;
    private readonly string _password;
    private readonly bool _debugMode;
    private readonly DiagnosticService _diagnosticService = new();
    private readonly StringBuilder _eventLog = new();
    private readonly DebugLogService _debugLog;

    private RdpAxHost? _rdpControl;
    private dynamic? _rdp;
    private DispatcherTimer? _statusWatchdog;
    private SessionToolbarWindow? _toolbarWindow;
    private bool _isCleaningUp;
    private bool _closeAfterDisconnect;
    private bool _isHostedFullScreen;

    public SessionWindow(HostEntry entry, string password, bool debugMode)
    {
        InitializeComponent();
        _entry = CloneEntry(entry);
        _password = password;
        _debugMode = debugMode;
        _debugLog = new DebugLogService($"session_{_entry.Name}");

        ShowInTaskbar = true;
        HostTitleText.Text = $"{_entry.Name} ({_entry.Address})";
        StatusText.Text = "准备连接...";
        Title = $"RDP 会话 - {_entry.Name}";

        if (_debugMode)
        {
            LogPathText.Text = $"调试日志: {_debugLog.LogFilePath}";
            LogPathText.Visibility = Visibility.Visible;
        }
        else
        {
            ConfigureMinimalPresentation();
        }

        Loaded += SessionWindow_Loaded;
        Activated += SessionWindow_Activated;
        StateChanged += SessionWindow_StateChanged;
        Closed += (_, _) => CleanupSession();
    }

    private void ConfigureMinimalPresentation()
    {
        HeaderPanel.Visibility = Visibility.Collapsed;
        FooterPanel.Visibility = Visibility.Collapsed;
        HostPanel.Margin = new Thickness(0);
        HostPanel.CornerRadius = new CornerRadius(0);
        HostPanel.BorderThickness = new Thickness(0);
        Background = System.Windows.Media.Brushes.Black;

        if (_entry.FullScreen)
        {
            _isHostedFullScreen = true;
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
        }
    }

    private void SessionWindow_Loaded(object sender, RoutedEventArgs e)
    {
        AppendEvent($"Window loaded. DebugMode={_debugMode}, FullScreen={_entry.FullScreen}");

        if (_entry.FullScreen)
        {
            WindowState = WindowState.Maximized;
            AppendEvent("Window maximized for full-screen mode.");
        }

        Activate();
        Topmost = true;
        Topmost = false;
        Focus();
        AppendEvent("Window activated and brought to front.");

        UpdateToolbarWindowVisibility();
        ConnectToHost();
    }

    private void SessionWindow_Activated(object? sender, EventArgs e)
    {
        AppendEvent("Window activated");
        UpdateToolbarWindowVisibility();
    }

    private void SessionWindow_StateChanged(object? sender, EventArgs e)
    {
        AppendEvent($"Window state changed to {WindowState}.");
        UpdateToolbarWindowVisibility();
    }

    private void ConnectToHost()
    {
        AppendEvent("ConnectToHost entered.");
        EnsureRdpControl();
        if (_rdp == null)
        {
            AppendEvent("ConnectToHost aborted because RDP control was not created.");
            return;
        }

        var (host, port) = SplitHostPort(_entry.Address, 3389);
        AppendEvent($"Host parsed to '{host}', port={port}.");

        if (string.IsNullOrWhiteSpace(_password))
        {
            StatusText.Text = "未找到密码，请返回主窗口重新输入。";
            AppendEvent("ConnectToHost aborted because password was empty.");
            return;
        }

        try
        {
            PlaceholderText.Visibility = Visibility.Collapsed;
            StatusText.Text = $"正在连接 {host}:{port}...";
            AppendEvent("Applying RDP connection properties.");

            _rdp.Server = host;
            _rdp.UserName = _entry.Username;
            _rdp.Domain = _entry.Domain;
            _rdp.ColorDepth = 32;

            var screen = Screen.PrimaryScreen;
            if (_entry.FullScreen && screen != null)
            {
                // Keep the RDP control inside our own borderless host window. A separate
                // toolbar window provides controls without shrinking the render area.
                _rdp.FullScreen = false;
                _rdp.DesktopWidth = screen.Bounds.Width;
                _rdp.DesktopHeight = screen.Bounds.Height;
                AppendEvent($"Configured hosted full-screen session {screen.Bounds.Width}x{screen.Bounds.Height}.");
            }
            else
            {
                _rdp.FullScreen = false;
                _rdp.DesktopWidth = _entry.DesktopWidth;
                _rdp.DesktopHeight = _entry.DesktopHeight;
                AppendEvent($"Configured windowed session {_entry.DesktopWidth}x{_entry.DesktopHeight}.");
            }

            if (TryGetAdvancedSettings(out var adv))
            {
                adv.RDPPort = port;
                adv.RedirectClipboard = _entry.RedirectClipboard;
                adv.SmartSizing = true;
                adv.EnableCredSspSupport = true;
                adv.NegotiateSecurityLayer = true;
                adv.AuthenticationLevel = 0;
                adv.EnableAutoReconnect = true;
                adv.ConnectToServerConsole = false;
                adv.DisplayConnectionBar = false;
                TrySet(adv, "ContainerHandledFullScreen", 1);
                TrySet(adv, "PromptForCredentials", false);
                TrySet(adv, "PromptForCredentialsOnClient", false);
                AppendEvent("Advanced settings applied.");
            }
            else
            {
                AppendEvent("AdvancedSettings9 unavailable.");
            }

            var ocx = _rdpControl?.OcxInstance;
            if (ocx != null)
            {
                TrySet(ocx, "PromptForCredentials", false);
                TrySet(ocx, "PromptForCredentialsOnClient", false);
                AppendEvent("OCX credential prompt suppression attempted.");
            }

            SetClearTextPassword(_rdp, _password);
            AppendEvent("Password assigned; invoking Connect().");
            _rdp.Connect();
            AppendEvent("Connect() returned without exception.");
            StartStatusWatchdog();
        }
        catch (Exception ex)
        {
            StatusText.Text = $"连接失败: {ex.Message}";
            PlaceholderText.Visibility = Visibility.Visible;
            AppendEvent($"Connect exception: {ex}");
            _ = RunDiagnosticsAsync(_entry.Address);
        }
    }

    private void EnsureRdpControl()
    {
        if (_rdp != null)
        {
            AppendEvent("EnsureRdpControl skipped because control already exists.");
            return;
        }

        var progIds = new[]
        {
            "MsRdpClient10NotSafeForScripting",
            "MsRdpClient9NotSafeForScripting",
            "MsRdpClient8NotSafeForScripting",
            "MsRdpClient7NotSafeForScripting",
            "MsRdpClient6NotSafeForScripting",
            "MsTscAx.MsTscAx"
        };

        Type? selectedType = null;
        string? selectedProgId = null;
        foreach (var progId in progIds)
        {
            var type = Type.GetTypeFromProgID(progId);
            AppendEvent($"Probing ActiveX ProgID: {progId} -> {(type == null ? "missing" : "found")}");
            if (type != null)
            {
                selectedType = type;
                selectedProgId = progId;
                break;
            }
        }

        if (selectedType == null)
        {
            StatusText.Text = "未找到系统 RDP ActiveX 控件。";
            DiagnosticText.Text = "请确认系统已启用远程桌面连接组件。";
            AppendEvent("No compatible RDP ActiveX control found.");
            return;
        }

        try
        {
            var clsid = selectedType.GUID.ToString("B");
            _rdpControl = new RdpAxHost(clsid);
            _rdpControl.BeginInit();
            RdpHost.Child = _rdpControl;
            _rdpControl.EndInit();

            _rdp = _rdpControl.OcxInstance;
            AppendEvent($"RDP ActiveX created. ProgID={selectedProgId}, CLSID={clsid}");
            WireRdpEvents();
        }
        catch (Exception ex)
        {
            StatusText.Text = $"创建 RDP 控件失败: {ex.Message}";
            AppendEvent($"EnsureRdpControl exception: {ex}");
        }
    }

    private void EnsureToolbarWindow()
    {
        if (_debugMode || !_isHostedFullScreen || _toolbarWindow != null)
        {
            return;
        }

        _toolbarWindow = new SessionToolbarWindow(_entry.Name)
        {
            Owner = this
        };
        _toolbarWindow.MinimizeRequested += (_, _) => MinimizeSessionWindow();
        _toolbarWindow.ExitFullScreenRequested += (_, _) => ExitHostedFullScreenMode();
        _toolbarWindow.DisconnectRequested += (_, _) => DisconnectSession();
        PositionToolbarWindow();
    }

    private void UpdateToolbarWindowVisibility()
    {
        if (_debugMode || !_isHostedFullScreen || !IsLoaded)
        {
            CloseToolbarWindow();
            return;
        }

        if (WindowState == WindowState.Minimized)
        {
            if (_toolbarWindow?.IsVisible == true)
            {
                _toolbarWindow.Hide();
            }
            return;
        }

        EnsureToolbarWindow();
        if (_toolbarWindow == null)
        {
            return;
        }

        if (!_toolbarWindow.IsVisible)
        {
            _toolbarWindow.Show();
        }
    }

    private void PositionToolbarWindow()
    {
        if (_toolbarWindow == null)
        {
            return;
        }

        _toolbarWindow.Left = Left + Math.Max(24, ActualWidth - _toolbarWindow.Width - 40);
        _toolbarWindow.Top = Top + 24;
    }

    private void CloseToolbarWindow()
    {
        if (_toolbarWindow == null)
        {
            return;
        }

        try
        {
            _toolbarWindow.Close();
        }
        catch
        {
            // ignore
        }
        finally
        {
            _toolbarWindow = null;
        }
    }

    private void Disconnect_Click(object sender, RoutedEventArgs e)
    {
        AppendEvent("Disconnect button clicked.");
        DisconnectSession();
    }

    private void MinimizeSessionWindow()
    {
        AppendEvent("Toolbar requested minimize.");
        WindowState = WindowState.Minimized;
        if (_toolbarWindow?.IsVisible == true)
        {
            _toolbarWindow.Hide();
        }
    }

    private void DisconnectSession()
    {
        if (_rdp == null)
        {
            AppendEvent("DisconnectSession: control already null, closing window.");
            Close();
            return;
        }

        try
        {
            var connected = (int)_rdp.Connected;
            AppendEvent($"DisconnectSession observed Connected={connected}.");
            if (connected != 0)
            {
                _closeAfterDisconnect = true;
                StatusText.Text = "正在断开...";
                _rdp.Disconnect();
                AppendEvent("Disconnect() invoked.");
            }
            else
            {
                AppendEvent("Session already disconnected, closing window.");
                Close();
            }
        }
        catch (Exception ex)
        {
            AppendEvent($"Disconnect exception: {ex}");
            CleanupSession();
            Close();
        }
    }

    private void CleanupSession()
    {
        if (_isCleaningUp)
        {
            return;
        }

        _isCleaningUp = true;
        AppendEvent("CleanupSession entered.");
        CloseToolbarWindow();
        _statusWatchdog?.Stop();
        _statusWatchdog = null;

        try
        {
            if (_rdp != null)
            {
                try
                {
                    var connected = (int)_rdp.Connected;
                    AppendEvent($"CleanupSession observed Connected={connected}.");
                    if (connected != 0)
                    {
                        _rdp.Disconnect();
                        AppendEvent("CleanupSession invoked Disconnect().");
                    }
                }
                catch (Exception ex)
                {
                    AppendEvent($"Cleanup disconnect check failed: {ex.Message}");
                }
            }

            if (_rdpControl != null)
            {
                RdpHost.Child = null;
                _rdpControl.Dispose();
                _rdpControl = null;
                AppendEvent("RDP host disposed.");
            }

            if (_rdp != null && Marshal.IsComObject(_rdp))
            {
                Marshal.FinalReleaseComObject(_rdp);
                AppendEvent("COM object released.");
            }
        }
        catch (Exception ex)
        {
            AppendEvent($"CleanupSession exception: {ex}");
        }
        finally
        {
            _rdp = null;
            _isCleaningUp = false;
        }
    }

    private void ExitHostedFullScreenMode()
    {
        if (!_isHostedFullScreen)
        {
            return;
        }

        _isHostedFullScreen = false;
        CloseToolbarWindow();
        WindowStyle = WindowStyle.SingleBorderWindow;
        ResizeMode = ResizeMode.CanResize;
        WindowState = WindowState.Normal;

        Width = Math.Max(_entry.DesktopWidth, 960);
        Height = Math.Max(_entry.DesktopHeight, 640);

        var workArea = SystemParameters.WorkArea;
        Left = workArea.Left + Math.Max(0, (workArea.Width - Width) / 2);
        Top = workArea.Top + Math.Max(0, (workArea.Height - Height) / 2);

        AppendEvent($"Exited hosted full-screen mode to windowed size {Width}x{Height}.");
    }

    private void WireRdpEvents()
    {
        if (_rdp == null)
        {
            return;
        }

        AttachEvent(_rdp, "OnConnected", nameof(Rdp_OnConnected));
        AttachEvent(_rdp, "OnDisconnected", nameof(Rdp_OnDisconnected));
        AttachEvent(_rdp, "OnFatalError", nameof(Rdp_OnFatalError));
        AttachEvent(_rdp, "OnLogonError", nameof(Rdp_OnLogonError));
        AttachEvent(_rdp, "OnWarning", nameof(Rdp_OnWarning));
        AttachEvent(_rdp, "OnConnecting", nameof(Rdp_OnConnecting));
        AttachEvent(_rdp, "OnLoginComplete", nameof(Rdp_OnLoginComplete));
        AppendEvent("RDP events wired.");
    }

    private void AttachEvent(object source, string eventName, string handlerName)
    {
        var evt = source.GetType().GetEvent(eventName);
        if (evt == null)
        {
            AppendEvent($"Event '{eventName}' not available on control.");
            return;
        }

        var method = GetType().GetMethod(handlerName, BindingFlags.Instance | BindingFlags.NonPublic);
        if (method == null)
        {
            AppendEvent($"Handler '{handlerName}' not found.");
            return;
        }

        try
        {
            var del = Delegate.CreateDelegate(evt.EventHandlerType!, this, method, throwOnBindFailure: false);
            if (del != null)
            {
                evt.AddEventHandler(source, del);
                AppendEvent($"Attached event '{eventName}'.");
            }
            else
            {
                AppendEvent($"Failed to create delegate for '{eventName}'.");
            }
        }
        catch (Exception ex)
        {
            AppendEvent($"AttachEvent exception for '{eventName}': {ex.Message}");
        }
    }

    private void Rdp_OnConnected(object? sender, EventArgs? e)
    {
        Dispatcher.Invoke(() =>
        {
            _statusWatchdog?.Stop();
            StatusText.Text = "已连接。";
            PlaceholderText.Visibility = Visibility.Collapsed;
            AppendEvent("RDP event: OnConnected");
        });
    }

    private void Rdp_OnDisconnected(object? sender, dynamic? e)
    {
        var reason = e?.discReason;
        var desc = GetDisconnectDescription(reason);
        Dispatcher.Invoke(() =>
        {
            _statusWatchdog?.Stop();
            StatusText.Text = $"已断开 (Code {reason})";
            DiagnosticText.Text = string.IsNullOrWhiteSpace(desc) ? "连接已断开。" : desc;
            PlaceholderText.Visibility = Visibility.Visible;
            AppendEvent($"RDP event: OnDisconnected reason={reason}, desc={desc}");
            CleanupSession();
            if (_closeAfterDisconnect)
            {
                Close();
            }
        });
    }

    private void Rdp_OnFatalError(object? sender, dynamic? e)
    {
        var code = e?.errorCode;
        Dispatcher.Invoke(() =>
        {
            StatusText.Text = $"致命错误: {code}";
            AppendEvent($"RDP event: OnFatalError code={code}");
        });
    }

    private void Rdp_OnLogonError(object? sender, dynamic? e)
    {
        var code = e?.lError;
        Dispatcher.Invoke(() =>
        {
            StatusText.Text = $"登录错误: {code}";
            AppendEvent($"RDP event: OnLogonError code={code}");
        });
    }

    private void Rdp_OnWarning(object? sender, dynamic? e)
    {
        var code = e?.warningCode;
        Dispatcher.Invoke(() =>
        {
            StatusText.Text = $"警告: {code}";
            AppendEvent($"RDP event: OnWarning code={code}");
        });
    }

    private void Rdp_OnConnecting(object? sender, EventArgs? e)
    {
        Dispatcher.Invoke(() => AppendEvent("RDP event: OnConnecting"));
    }

    private void Rdp_OnLoginComplete(object? sender, EventArgs? e)
    {
        Dispatcher.Invoke(() => AppendEvent("RDP event: OnLoginComplete"));
    }

    private void StartStatusWatchdog()
    {
        _statusWatchdog?.Stop();
        _statusWatchdog = new DispatcherTimer { Interval = TimeSpan.FromSeconds(15) };
        _statusWatchdog.Tick += (_, _) =>
        {
            _statusWatchdog?.Stop();
            if (StatusText.Text.StartsWith("正在连接"))
            {
                StatusText.Text = "连接超时或无响应，请检查网络、凭据和远程桌面服务。";
                AppendEvent("Status watchdog fired after 15 seconds.");
            }
        };
        _statusWatchdog.Start();
        AppendEvent("Status watchdog started.");
    }

    private async Task RunDiagnosticsAsync(string host)
    {
        if (string.IsNullOrWhiteSpace(host))
        {
            StatusText.Text = "诊断失败：地址为空。";
            AppendEvent("RunDiagnosticsAsync aborted because host was empty.");
            return;
        }

        StatusText.Text = "正在诊断网络...";
        DiagnosticText.Text = string.Empty;
        var (parsedHost, port) = SplitHostPort(host, 3389);
        AppendEvent($"Running diagnostics for {parsedHost}:{port}.");
        var result = await Task.Run(() => _diagnosticService.Run(parsedHost, port));
        Dispatcher.Invoke(() =>
        {
            DiagnosticText.Text = string.Join(Environment.NewLine, result.Lines);
            AppendEvent($"Diagnostics completed. Success={result.Success}");
            if (!result.Success)
            {
                PlaceholderText.Visibility = Visibility.Visible;
            }
        });
    }

    private bool TryGetAdvancedSettings(out dynamic adv)
    {
        adv = null!;
        if (_rdp == null)
        {
            return false;
        }

        try
        {
            adv = _rdp.AdvancedSettings9;
            return adv != null;
        }
        catch (Exception ex)
        {
            AppendEvent($"AdvancedSettings9 access failed: {ex.Message}");
            return false;
        }
    }

    private void SetClearTextPassword(dynamic rdp, string password)
    {
        try
        {
            rdp.GetType().InvokeMember("ClearTextPassword", BindingFlags.SetProperty, null, rdp, new object[] { password });
            AppendEvent("ClearTextPassword set on main RDP object.");
            return;
        }
        catch (Exception ex)
        {
            AppendEvent($"ClearTextPassword on main object failed: {ex.Message}");
        }

        try
        {
            if (TryGetAdvancedSettings(out var adv))
            {
                adv.GetType().InvokeMember("ClearTextPassword", BindingFlags.SetProperty, null, adv, new object[] { password });
                AppendEvent("ClearTextPassword set on advanced settings.");
            }
        }
        catch (Exception ex)
        {
            AppendEvent($"ClearTextPassword on advanced settings failed: {ex.Message}");
        }
    }

    private (string host, int port) SplitHostPort(string input, int defaultPort)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return (string.Empty, defaultPort);
        }

        var trimmed = input.Trim();

        if (trimmed.StartsWith("[") && trimmed.Contains("]"))
        {
            var end = trimmed.IndexOf(']');
            var hostPart = trimmed.Substring(1, end - 1);
            var remaining = trimmed[(end + 1)..];
            if (remaining.StartsWith(":") && int.TryParse(remaining[1..], out var port))
            {
                return (hostPart, port);
            }

            return (hostPart, defaultPort);
        }

        var lastColon = trimmed.LastIndexOf(':');
        if (lastColon > 0 &&
            lastColon < trimmed.Length - 1 &&
            int.TryParse(trimmed[(lastColon + 1)..], out var parsedPort))
        {
            return (trimmed[..lastColon], parsedPort);
        }

        return (trimmed, defaultPort);
    }

    private string GetDisconnectDescription(int? reason)
    {
        if (_rdp == null)
        {
            return string.Empty;
        }

        try
        {
            var ext = 0u;
            try
            {
                ext = (uint)(_rdp.ExtendedDisconnectReason ?? 0);
            }
            catch (Exception ex)
            {
                AppendEvent($"ExtendedDisconnectReason read failed: {ex.Message}");
            }

            var desc = _rdp.GetErrorDescription((uint)(reason ?? 0), ext);
            return desc is string text ? text : string.Empty;
        }
        catch (Exception ex)
        {
            AppendEvent($"GetDisconnectDescription failed: {ex.Message}");
            return string.Empty;
        }
    }

    private void AppendEvent(string message)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {message}";
        _eventLog.AppendLine(line);
        if (_debugMode)
        {
            EventLogText.Text = _eventLog.ToString();
            EventLogText.ScrollToEnd();
        }
        _debugLog.Write("SessionWindow", message);
    }

    private void TrySet(object target, string propertyName, object value)
    {
        try
        {
            target.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, target, new[] { value });
        }
        catch (Exception ex)
        {
            AppendEvent($"TrySet failed for '{propertyName}': {ex.Message}");
        }
    }

    private static HostEntry CloneEntry(HostEntry entry)
    {
        return new HostEntry
        {
            Name = entry.Name,
            Address = entry.Address,
            Username = entry.Username,
            Domain = entry.Domain,
            DesktopWidth = entry.DesktopWidth,
            DesktopHeight = entry.DesktopHeight,
            RedirectClipboard = entry.RedirectClipboard,
            FullScreen = entry.FullScreen
        };
    }
}
