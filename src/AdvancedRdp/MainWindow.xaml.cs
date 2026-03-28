using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using AdvancedRdp.Models;
using AdvancedRdp.Services;

namespace AdvancedRdp;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private readonly CredentialService _credentialService;
    private readonly HostStore _hostStore;
    private readonly DebugLogService _debugLog = new("main");

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
        DebugModeCheck.Checked += (_, _) => UpdateDebugHint();
        DebugModeCheck.Unchecked += (_, _) => UpdateDebugHint();
        LoadHosts();
        UpdateDebugHint();
        WriteDebug("Main window initialized.");
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

        WriteDebug($"Loaded {Hosts.Count} host entries.");
    }

    private void OnPropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    private void SaveHost_Click(object sender, RoutedEventArgs e)
    {
        var entry = ReadForm();
        if (entry is null)
        {
            return;
        }

        var existing = SelectedHost;
        if (existing == null)
        {
            Hosts.Add(entry);
            SelectedHost = entry;
            WriteDebug($"Created host entry '{entry.Name}' -> {entry.Address}.");
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
            existing.FullScreen = entry.FullScreen;
            HostList.Items.Refresh();
            WriteDebug($"Updated host entry '{entry.Name}' -> {entry.Address}.");
        }

        _hostStore.SaveAll(Hosts);

        if (SavePasswordCheck.IsChecked == true)
        {
            var password = PasswordInput.Password;
            if (!string.IsNullOrEmpty(password))
            {
                _credentialService.SavePassword(entry.CredentialKey, password);
                WriteDebug($"Saved password for '{entry.Name}'.");
            }
        }

        StatusText.Text = "配置已保存。";
    }

    private void Connect_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var entry = ReadForm();
            if (entry is null)
            {
                return;
            }

            var debugMode = DebugModeCheck.IsChecked == true;
            WriteDebug($"Connect clicked. Host='{entry.Name}', Address='{entry.Address}', User='{entry.Username}', FullScreen={entry.FullScreen}, DebugMode={debugMode}.");
            StatusText.Text = $"已点击连接，正在准备会话: {entry.Name}";

            if (SavePasswordCheck.IsChecked == true && !string.IsNullOrWhiteSpace(PasswordInput.Password))
            {
                _credentialService.SavePassword(entry.CredentialKey, PasswordInput.Password);
                WriteDebug($"Connect path saved password for '{entry.Name}'.");
            }

            var password = string.IsNullOrWhiteSpace(PasswordInput.Password)
                ? _credentialService.GetPassword(entry.CredentialKey)
                : PasswordInput.Password;

            if (string.IsNullOrWhiteSpace(password))
            {
                StatusText.Text = "未找到密码，请输入后重试。";
                WriteDebug("Connect aborted because password was empty.");
                return;
            }

            var sessionWindow = new SessionWindow(entry, password, debugMode);
            sessionWindow.Show();
            sessionWindow.Activate();
            StatusText.Text = $"会话窗口已打开: {entry.Name}";

            if (debugMode)
            {
                DebugHintText.Text = $"调试模式已开启。主日志: {_debugLog.LogFilePath}";
            }

            WriteDebug($"Session window shown for '{entry.Name}'.");
        }
        catch (Exception ex)
        {
            StatusText.Text = $"打开会话窗口失败: {ex.Message}";
            WriteDebug($"Connect_Click exception: {ex}");
            if (DebugModeCheck.IsChecked == true)
            {
                System.Windows.MessageBox.Show(this, ex.ToString(), "连接失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void DeleteHost_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedHost == null)
        {
            return;
        }

        WriteDebug($"Deleting host entry '{SelectedHost.Name}'.");
        _credentialService.DeletePassword(SelectedHost.CredentialKey);
        Hosts.Remove(SelectedHost);
        SelectedHost = null;
        _hostStore.SaveAll(Hosts);
        ClearForm();
        StatusText.Text = "主机已删除。";
    }

    private void NewHost_Click(object sender, RoutedEventArgs e)
    {
        HostList.SelectedIndex = -1;
        SelectedHost = null;
        ClearForm();
        StatusText.Text = "已创建空白主机配置。";
        WriteDebug("Prepared new host form.");
    }

    private void HostList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (HostList.SelectedItem is HostEntry entry)
        {
            SelectedHost = entry;
            ApplySelectionToForm(entry);
            WriteDebug($"Selected host '{entry.Name}'.");
        }
    }

    private HostEntry? ReadForm()
    {
        if (string.IsNullOrWhiteSpace(NameInput.Text) || string.IsNullOrWhiteSpace(AddressInput.Text))
        {
            StatusText.Text = "名称和地址不能为空。";
            WriteDebug("ReadForm failed because name or address was empty.");
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

    private void TryFillSavedPassword(HostEntry entry)
    {
        try
        {
            var saved = _credentialService.GetPassword(entry.CredentialKey);
            if (!string.IsNullOrEmpty(saved))
            {
                PasswordInput.Password = saved;
                WriteDebug($"Loaded saved password for '{entry.Name}'.");
            }
        }
        catch (Exception ex)
        {
            WriteDebug($"Failed to load saved password for '{entry.Name}': {ex.Message}");
        }
    }

    private void UpdateDebugHint()
    {
        if (DebugModeCheck.IsChecked == true)
        {
            DebugHintText.Text = $"调试模式已开启。主日志文件: {_debugLog.LogFilePath}";
        }
        else
        {
            DebugHintText.Text = string.Empty;
        }
    }

    private void WriteDebug(string message)
    {
        _debugLog.Write("MainWindow", message);
    }
}
