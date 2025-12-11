using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Reflection;
using System.Windows;
using AdvancedRdp.Controls;
using AdvancedRdp.Models;
using AdvancedRdp.Services;

namespace AdvancedRdp;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private readonly CredentialService _credentialService;
    private readonly HostStore _hostStore;
    private RdpAxHost? _rdpControl;
    private dynamic? _rdp;

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
            RedirectClipboard = RedirectClipboardCheck.IsChecked == true
        };
    }

    private void ApplySelectionToForm(HostEntry entry)
    {
        NameInput.Text = entry.Name;
        AddressInput.Text = entry.Address;
        UserInput.Text = entry.Username;
        DomainInput.Text = entry.Domain;
        RedirectClipboardCheck.IsChecked = entry.RedirectClipboard;
        PasswordInput.Password = string.Empty;
    }

    private void ClearForm()
    {
        NameInput.Text = string.Empty;
        AddressInput.Text = string.Empty;
        UserInput.Text = string.Empty;
        DomainInput.Text = string.Empty;
        RedirectClipboardCheck.IsChecked = true;
        PasswordInput.Password = string.Empty;
    }

    private void ConnectToHost(HostEntry entry)
    {
        EnsureRdpControl();
        if (_rdp == null) return;

        var password = string.IsNullOrWhiteSpace(PasswordInput.Password)
            ? _credentialService.GetPassword(entry.CredentialKey)
            : PasswordInput.Password;

        if (string.IsNullOrEmpty(password))
        {
            StatusText.Text = "未找到密码，请输入后再试";
            return;
        }

        try
        {
            PlaceholderText.Visibility = Visibility.Collapsed;
            StatusText.Text = $"正在连接 {entry.Address}...";

            _rdp.Server = entry.Address;
            _rdp.UserName = entry.Username;
            _rdp.Domain = entry.Domain;
            _rdp.ColorDepth = 32;
            _rdp.DesktopWidth = entry.DesktopWidth;
            _rdp.DesktopHeight = entry.DesktopHeight;

            if (TryGetAdvancedSettings(out var adv))
            {
                adv.RDPPort = 3389;
                adv.RedirectClipboard = entry.RedirectClipboard;
                adv.SmartSizing = true;
            }

            SetClearTextPassword(_rdp, password);
            _rdp.Connect();
            StatusText.Text = "正在连接...";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"连接失败: {ex.Message}";
            PlaceholderText.Visibility = Visibility.Visible;
        }
    }

    private void EnsureRdpControl()
    {
        if (_rdp != null) return;

        var type = Type.GetTypeFromProgID("MsRdpClient9NotSafeForScripting");
        if (type == null)
        {
            StatusText.Text = "未找到 RDP ActiveX 控件";
            return;
        }

        var clsid = type.GUID.ToString("B");
        _rdpControl = new RdpAxHost(clsid);
        _rdpControl.BeginInit();
        RdpHost.Child = _rdpControl;
        _rdpControl.EndInit();

        _rdp = _rdpControl.OcxInstance;
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
}
