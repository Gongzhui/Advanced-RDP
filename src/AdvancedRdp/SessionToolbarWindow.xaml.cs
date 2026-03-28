using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace AdvancedRdp;

public partial class SessionToolbarWindow : Window
{
    public event EventHandler? MinimizeRequested;
    public event EventHandler? ExitFullScreenRequested;
    public event EventHandler? DisconnectRequested;

    public SessionToolbarWindow(string title)
    {
        InitializeComponent();
        TitleText.Text = title;
        PreviewMouseLeftButtonDown += SessionToolbarWindow_PreviewMouseLeftButtonDown;
    }

    private void SessionToolbarWindow_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.OriginalSource is DependencyObject source && FindAncestor<System.Windows.Controls.Button>(source) != null)
        {
            return;
        }

        try
        {
            DragMove();
        }
        catch
        {
            // ignore drag failures
        }
    }

    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
    {
        MinimizeRequested?.Invoke(this, EventArgs.Empty);
    }

    private void ExitFullScreenButton_Click(object sender, RoutedEventArgs e)
    {
        ExitFullScreenRequested?.Invoke(this, EventArgs.Empty);
    }

    private void DisconnectButton_Click(object sender, RoutedEventArgs e)
    {
        DisconnectRequested?.Invoke(this, EventArgs.Empty);
    }

    private static T? FindAncestor<T>(DependencyObject current) where T : DependencyObject
    {
        var node = current;
        while (node != null)
        {
            if (node is T match)
            {
                return match;
            }

            node = System.Windows.Media.VisualTreeHelper.GetParent(node);
        }

        return null;
    }
}
