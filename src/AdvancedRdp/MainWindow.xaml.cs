using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using System.Text;
using System.Windows.Forms;
using AdvancedRdp.Controls;
using AdvancedRdp.Models;
using AdvancedRdp.Services;

namespace AdvancedRdp;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private readonly CredentialService _credentialService;
    private readonly HostStore _hostStore;
    private readonly DiagnosticService _diagnosticService;
    private RdpAxHost? _rdpControl;
    private dynamic? _rdp;
    private readonly StringBuilder _eventLog = new();

    public ObservableCollection<HostEntry> Hosts { get; } = new();

    private HostEntry? _selectedHost;
    public HostEntry? SelectedHost
    {
        get => _selectedHost;
        set
        {
            _selectedHost = value;
            OnPropertyChanged(nameof(SelectedHost));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
        _credentialService = new CredentialService("AdvancedRdp");
        _hostStore = new HostStore();
        _diagnosticService = new DiagnosticService();

        LoadHosts();
        this.Closed += (_, _) => DisconnectIfNeeded();
    }

    private void LoadHosts()
    {
        Hosts.Clear();
        foreach (var host in _hostStore.LoadAll())
        {
            Hosts.Add(host);
        }

        if (Hosts.Count > 0)
        {
            HostList.SelectedIndex = 0;
            ApplySelectionToForm(Hosts[0]);
        }
    }

    private void OnPropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    private void SaveHost_Click(object sender, RoutedEventArgs e)
    {
        var entry = ReadForm();
        if (entry is null) return;

        var existing = SelectedHost;
        if (existing == null)
        {
            Hosts.Add(entry);
            SelectedHost = entry;
        }
        else
        {
            existing.Name = entry.Name;
            existing.Address = entry.Address;
            existing.Username = entry.Username;
            existing.Domain = entry.Domain;
            existing.DesktopWidth = entry.DesktopWidth;
            existing.DesktopHeight = entry.DesktopHeight;
            existing.RedirectClipboard = entry.RedirectClipboard;
            HostList.Items.Refresh();
        }

        _hostStore.SaveAll(Hosts);

        if (SavePasswordCheck.IsChecked == true)
        {
            var password = PasswordInput.Password;
            if (!string.IsNullOrEmpty(password))
            {
                _credentialService.SavePassword(entry.CredentialKey, password);
            }
        }

        StatusText.Text = "已保存配置";
    }

    private void Connect_Click(object sender, RoutedEventArgs e)
    {
        var entry = ReadForm();
        if (entry is null) return;

        if (SavePasswordCheck.IsChecked == true && !string.IsNullOrWhiteSpace(PasswordInput.Password))
        {
            _credentialService.SavePassword(entry.CredentialKey, PasswordInput.Password);
        }

        ConnectToHost(entry);
    }

    private void DeleteHost_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedHost == null) return;
        _credentialService.DeletePassword(SelectedHost.CredentialKey);
        Hosts.Remove(SelectedHost);
        SelectedHost = null;
        _hostStore.SaveAll(Hosts);
        ClearForm();
        StatusText.Text = "已删除主机";
    }

    private void NewHost_Click(object sender, RoutedEventArgs e)
    {
        HostList.SelectedIndex = -1;
        SelectedHost = null;
        ClearForm();
        StatusText.Text = "已创建空白主机";
    }

    private void HostList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (HostList.SelectedItem is HostEntry entry)
        {
            SelectedHost = entry;
            ApplySelectionToForm(entry);
        }
    }

    private HostEntry? ReadForm()
    {
        if (string.IsNullOrWhiteSpace(NameInput.Text) || string.IsNullOrWhiteSpace(AddressInput.Text))
        {
            StatusText.Text = "名称和地址不能为空";
            return null;
        }

        return new HostEntry
        {
            Name = NameInput.Text.Trim(),
            Address = AddressInput.Text.Trim(),
            Username = UserInput.Text.Trim(),
            Domain = DomainInput.Text.Trim(),
            DesktopWidth = 1280,
            DesktopHeight = 720,
            RedirectClipboard = RedirectClipboardCheck.IsChecked == true,
            FullScreen = FullScreenCheck.IsChecked == true
        };
    }

    private void ApplySelectionToForm(HostEntry entry)
    {
        NameInput.Text = entry.Name;
        AddressInput.Text = entry.Address;
        UserInput.Text = entry.Username;
        DomainInput.Text = entry.Domain;
        RedirectClipboardCheck.IsChecked = entry.RedirectClipboard;
        FullScreenCheck.IsChecked = entry.FullScreen;
        PasswordInput.Password = string.Empty;
        TryFillSavedPassword(entry);
    }

    private void ClearForm()
    {
        NameInput.Text = string.Empty;
        AddressInput.Text = string.Empty;
        UserInput.Text = string.Empty;
        DomainInput.Text = string.Empty;
        RedirectClipboardCheck.IsChecked = true;
        FullScreenCheck.IsChecked = true;
        PasswordInput.Password = string.Empty;
    }

    private void ConnectToHost(HostEntry entry)
    {
        EnsureRdpControl();
        if (_rdp == null) return;

        var (host, port) = SplitHostPort(entry.Address, 3389);

        var password = string.IsNullOrWhiteSpace(PasswordInput.Password)
            ? _credentialService.GetPassword(entry.CredentialKey)
            : PasswordInput.Password;

        if (string.IsNullOrEmpty(password))
        {
            StatusText.Text = "未找到密码，请输入后重试";
            return;
        }

        try
        {
            PlaceholderText.Visibility = Visibility.Collapsed;
            StatusText.Text = $"正在连接 {entry.Address}...";

            _rdp.Server = host;
            _rdp.UserName = entry.Username;
            _rdp.Domain = entry.Domain;
            _rdp.ColorDepth = 32;
            if (entry.FullScreen)
            {
                var screen = Screen.PrimaryScreen;
                if (screen != null)
                {
                    _rdp.DesktopWidth = screen.Bounds.Width;
                    _rdp.DesktopHeight = screen.Bounds.Height;
                }
                _rdp.FullScreen = true;
            }
            else
            {
                _rdp.FullScreen = false;
                _rdp.DesktopWidth = entry.DesktopWidth;
                _rdp.DesktopHeight = entry.DesktopHeight;
            }

            if (TryGetAdvancedSettings(out var adv))
            {
                adv.RDPPort = port;
                adv.RedirectClipboard = entry.RedirectClipboard;
                adv.SmartSizing = true;
                adv.EnableCredSspSupport = true; // enable NLA/CredSSP
                adv.NegotiateSecurityLayer = true;
                adv.AuthenticationLevel = 0; // 0 = no auth prompt (ignore cert warnings)
                adv.EnableAutoReconnect = true;
                adv.ConnectToServerConsole = false;
                adv.DisplayConnectionBar = entry.FullScreen;
                TrySet(adv, "PromptForCredentials", false);
                TrySet(adv, "PromptForCredentialsOnClient", false);
            }

            // Best effort to silence cert warnings for non-scriptable layer as well
            var ocx = _rdpControl?.OcxInstance;
            if (ocx != null)
            {
                TrySet(ocx, "PromptForCredentials", false);
                TrySet(ocx, "PromptForCredentialsOnClient", false);
            }

            SetClearTextPassword(_rdp, password);
            _rdp.Connect();
            StatusText.Text = $"正在连接 {host}:{port}...";
            StartStatusWatchdog();
        }
        catch (Exception ex)
        {
            StatusText.Text = $"连接失败: {ex.Message}";
            PlaceholderText.Visibility = Visibility.Visible;
            _ = RunDiagnosticsAsync(entry.Address);
        }
    }

    private void EnsureRdpControl()
    {
        if (_rdp != null) return;

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
        foreach (var progId in progIds)
        {
            var type = Type.GetTypeFromProgID(progId);
            if (type != null)
            {
                selectedType = type;
                StatusText.Text = $"已找到 ActiveX: {progId}";
                break;
            }
        }

        if (selectedType == null)
        {
            StatusText.Text = "未找到 RDP ActiveX 控件（请启用系统远程桌面组件）";
            return;
        }

        var clsid = selectedType.GUID.ToString("B");
        _rdpControl = new RdpAxHost(clsid);
        _rdpControl.BeginInit();
        RdpHost.Child = _rdpControl;
        _rdpControl.EndInit();

        _rdp = _rdpControl.OcxInstance;
        WireRdpEvents();
    }

    private void DisconnectIfNeeded()
    {
        if (_rdp != null)
        {
            try
            {
                var connected = (int)_rdp.Connected;
                if (connected == 1)
                {
                    _rdp.Disconnect();
                }
            }
            catch
            {
                // ignore
            }
        }
    }

    private void WireRdpEvents()
    {
        if (_rdp == null) return;

        AttachEvent(_rdp, "OnConnected", nameof(Rdp_OnConnected));
        AttachEvent(_rdp, "OnDisconnected", nameof(Rdp_OnDisconnected));
        AttachEvent(_rdp, "OnFatalError", nameof(Rdp_OnFatalError));
        AttachEvent(_rdp, "OnLogonError", nameof(Rdp_OnLogonError));
        AttachEvent(_rdp, "OnWarning", nameof(Rdp_OnWarning));
        AttachEvent(_rdp, "OnConnecting", nameof(Rdp_OnConnecting));
        AttachEvent(_rdp, "OnLoginComplete", nameof(Rdp_OnLoginComplete));
    }

    private void AttachEvent(object source, string eventName, string handlerName)
    {
        var evt = source.GetType().GetEvent(eventName);
        if (evt == null) return;

        var method = GetType().GetMethod(handlerName, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
        if (method == null) return;

        try
        {
            var del = Delegate.CreateDelegate(evt.EventHandlerType!, this, method, throwOnBindFailure: false);
            if (del != null)
            {
                evt.AddEventHandler(source, del);
            }
        }
        catch
        {
            // ignore if cannot attach
        }
    }

    // Event handlers (using dynamic for compatibility with COM event args)
    private void Rdp_OnConnected(object? sender, EventArgs? e)
    {
        Dispatcher.Invoke(() =>
        {
            StatusText.Text = "已连接";
            PlaceholderText.Visibility = Visibility.Collapsed;
            AppendEvent("Connected");
        });
    }

    private void Rdp_OnDisconnected(object? sender, dynamic? e)
    {
        var reason = e?.discReason;
        var desc = GetDisconnectDescription(reason);
        Dispatcher.Invoke(() =>
        {
            StatusText.Text = $"已断开 (Code {reason})";
            if (!string.IsNullOrWhiteSpace(desc))
            {
                DiagnosticText.Text = desc;
            }
            PlaceholderText.Visibility = Visibility.Visible;
            AppendEvent($"Disconnected code={reason} desc={desc}");
        });
    }

    private void Rdp_OnFatalError(object? sender, dynamic? e)
    {
        var code = e?.errorCode;
        Dispatcher.Invoke(() =>
        {
            StatusText.Text = $"致命错误: {code}";
            AppendEvent($"FatalError code={code}");
        });
    }

    private void Rdp_OnLogonError(object? sender, dynamic? e)
    {
        var code = e?.lError;
        Dispatcher.Invoke(() =>
        {
            StatusText.Text = $"登录错误: {code}";
            AppendEvent($"LogonError code={code}");
        });
    }

    private void Rdp_OnWarning(object? sender, dynamic? e)
    {
        var code = e?.warningCode;
        Dispatcher.Invoke(() =>
        {
            StatusText.Text = $"警告: {code}";
            AppendEvent($"Warning code={code}");
        });
    }

    private void Rdp_OnConnecting(object? sender, EventArgs? e)
    {
        Dispatcher.Invoke(() => AppendEvent("Connecting"));
    }

    private void Rdp_OnLoginComplete(object? sender, EventArgs? e)
    {
        Dispatcher.Invoke(() => AppendEvent("LoginComplete"));
    }

    private void StartStatusWatchdog()
    {
        // If the control stays in "connecting" too long without events, we still surface a hint.
        var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(15) };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            if (StatusText.Text.StartsWith("正在连接"))
            {
                StatusText.Text = "连接超时或未响应，请检查网络/凭据/防火墙";
            }
        };
        timer.Start();
    }

    private async Task RunDiagnosticsAsync(string host)
    {
        if (string.IsNullOrWhiteSpace(host))
        {
            StatusText.Text = "诊断失败：地址为空";
            return;
        }

        StatusText.Text = "正在诊断网络...";
        DiagnosticText.Text = string.Empty;
        var (parsedHost, port) = SplitHostPort(host, 3389);
        var result = await Task.Run(() => _diagnosticService.Run(parsedHost, port));
        Dispatcher.Invoke(() =>
        {
            DiagnosticText.Text = string.Join(Environment.NewLine, result.Lines);
            StatusText.Text = result.Success
                ? $"诊断完成：{parsedHost}:{port} 可达"
                : $"诊断提示：{parsedHost}:{port} 网络/端口/凭据可能有问题";
            AppendEvent($"Diag host={parsedHost} port={port} success={result.Success}");
        });
    }

    private bool TryGetAdvancedSettings(out dynamic adv)
    {
        adv = null!;
        if (_rdp == null) return false;
        try
        {
            adv = _rdp.AdvancedSettings9;
            return adv != null;
        }
        catch
        {
            return false;
        }
    }

    private void SetClearTextPassword(dynamic rdp, string password)
    {
        try
        {
            rdp.GetType().InvokeMember("ClearTextPassword", BindingFlags.SetProperty, null, rdp, new object[] { password });
            return;
        }
        catch
        {
            // ignore and try advanced settings
        }

        try
        {
            if (TryGetAdvancedSettings(out var adv))
            {
                adv.GetType().InvokeMember("ClearTextPassword", BindingFlags.SetProperty, null, adv, new object[] { password });
            }
        }
        catch
        {
            // ignore
        }
    }

    private (string host, int port) SplitHostPort(string input, int defaultPort)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return (string.Empty, defaultPort);
        }

        var trimmed = input.Trim();

        // If IPv6 with [], handle [addr]:port
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
        if (lastColon > 0 && lastColon < trimmed.Length - 1 && int.TryParse(trimmed[(lastColon + 1)..], out var p))
        {
            return (trimmed[..lastColon], p);
        }

        return (trimmed, defaultPort);
    }

    private string GetDisconnectDescription(int? reason)
    {
        if (_rdp == null) return string.Empty;
        try
        {
            var ext = 0u;
            try
            {
                ext = (uint)(_rdp.ExtendedDisconnectReason ?? 0);
            }
            catch
            {
                // ignore
            }

            var desc = _rdp.GetErrorDescription((uint)(reason ?? 0), ext);
            return desc is string s ? s : string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private void AppendEvent(string message)
    {
        _eventLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] {message}");
        EventLogText.Text = _eventLog.ToString();
    }

    private void TrySet(object target, string propertyName, object value)
    {
        try
        {
            target.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, target, new[] { value });
        }
        catch
        {
            // ignore best-effort setters
        }
    }

    private void TryFillSavedPassword(HostEntry entry)
    {
        try
        {
            var saved = _credentialService.GetPassword(entry.CredentialKey);
            if (!string.IsNullOrEmpty(saved))
            {
                PasswordInput.Password = saved;
            }
        }
        catch
        {
            // ignore
        }
    }
}
