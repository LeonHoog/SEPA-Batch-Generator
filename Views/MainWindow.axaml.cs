using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Layout;
using Avalonia.Platform.Storage;
using SEPA_Batch_Generator.ViewModels;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace SEPA_Batch_Generator.Views;

public partial class MainWindow : Window
{
    private MainWindowViewModel? _viewModel;
    private static readonly Regex WindowsPathRegex = new(@"([A-Za-z]:\\.+)$", RegexOptions.Compiled);
    private static readonly Regex UncPathRegex = new(@"(\\\\[^\s].+)$", RegexOptions.Compiled);

    public MainWindow()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
    }

    private void OnDataContextChanged(object? sender, EventArgs e)
    {
        _viewModel?.PropertyChanged -= OnViewModelPropertyChanged;

        _viewModel = DataContext as MainWindowViewModel;
        _viewModel?.PropertyChanged += OnViewModelPropertyChanged;
    }

    private async void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_viewModel is null || e.PropertyName != nameof(MainWindowViewModel.PendingWarningMessage))
            return;

        if (string.IsNullOrWhiteSpace(_viewModel.PendingWarningMessage))
            return;

        string message = _viewModel.PendingWarningMessage;
        _viewModel.PendingWarningMessage = string.Empty;
        await ShowWarningDialogAsync("Waarschuwing", message);
    }

    private async void BrowseExcelFile_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        IReadOnlyList<IStorageFile> selected = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Selecteer Excel bestand",
            AllowMultiple = false,
            FileTypeFilter =
                [
                    new("Excel bestanden")
                    {
                        Patterns = ["*.xlsx", "*.xlsm", "*.xls"]
                    }
                ]
        });

        IStorageFile? file = selected.Count > 0 ? selected[0] : null;
        string? localPath = file?.TryGetLocalPath();

        if (!string.IsNullOrWhiteSpace(localPath) && DataContext is MainWindowViewModel vm)
            vm.ExcelPath = localPath;
    }

    private async void BrowseOutputFolder_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        IReadOnlyList<IStorageFolder> folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "Selecteer output map",
            AllowMultiple = false
        });

        IStorageFolder? folder = folders.Count > 0 ? folders[0] : null;
        string? localPath = folder?.TryGetLocalPath();

        if (!string.IsNullOrWhiteSpace(localPath) && DataContext is MainWindowViewModel vm)
        {
            vm.OutputFolder = localPath;
        }
    }

    private async void BrowseLogFile_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm)
            return;

        string? suggestedName = string.IsNullOrWhiteSpace(vm.LogFilePath)
            ? "sepa-log.txt"
            : Path.GetFileName(vm.LogFilePath);

        IStorageFile? file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Kies logbestand",
            SuggestedFileName = string.IsNullOrWhiteSpace(suggestedName) ? "sepa-log.txt" : suggestedName,
            DefaultExtension = "txt",
            FileTypeChoices =
                [
                    new("Tekstbestand")
                    {
                        Patterns = ["*.txt"],
                        MimeTypes = ["text/plain"]
                    }
                ]
        });

        string? localPath = file?.TryGetLocalPath();
        if (!string.IsNullOrWhiteSpace(localPath))
        {
            vm.LogFilePath = localPath;
        }
    }

    private async Task ShowWarningDialogAsync(string title, string message)
    {
        Button okButton = new()
        {
            Content = "OK",
            HorizontalAlignment = HorizontalAlignment.Center,
            MinWidth = 90
        };

        StackPanel panel = new()
        {
            Margin = new Thickness(16),
            Spacing = 12,
            Children =
                {
                    new TextBlock { Text = message, TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                    okButton
                }
        };

        Window dialog = new()
        {
            Title = title,
            Width = 460,
            Height = 180,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Content = panel
        };

        okButton.Click += (_, _) => dialog.Close();
        await dialog.ShowDialog(this);
    }

    private async void MessagesListBox_DoubleTapped(object? sender, TappedEventArgs e)
    {
        if (sender is not ListBox listBox)
            return;

        string? message = listBox.SelectedItem as string;
        if (string.IsNullOrWhiteSpace(message) && e.Source is Control control)
            message = control.DataContext as string;

        if (string.IsNullOrWhiteSpace(message))
            return;

        // Check if this is the "Totaalbedrag" message
        if (message.Contains("Totaalbedrag") && _viewModel is not null)
        {
            string breakdown = _viewModel.GetAmountBreakdown();
            await ShowBreakdownDialog(breakdown);
            return;
        }

        string? path = TryExtractPathFromMessage(message);
        if (string.IsNullOrWhiteSpace(path))
            return;

        if (File.Exists(path) || Directory.Exists(path))
            await Clipboard!.SetTextAsync(path);
    }

    private async Task ShowBreakdownDialog(string breakdown)
    {
        TextBlock textBlock = new()
        {
            Text = breakdown,
            FontFamily = new("Courier New"),
            FontSize = 12,
            Foreground = Foreground,
            Margin = new Thickness(0, 0, 0, 0)
        };

        ScrollViewer scrollViewer = new()
        {
            Content = textBlock,
            Height = 400
        };

        Button okButton = new()
        {
            Content = "OK",
            HorizontalAlignment = HorizontalAlignment.Right,
            MinWidth = 100
        };

        StackPanel panel = new()
        {
            Spacing = 12,
            Children = { scrollViewer, okButton }
        };

        Window dialog = new()
        {
            Content = panel,
            Width = 700,
            Height = 500,
            Title = "Overboeking overzicht",
            CanResize = true,
            ShowInTaskbar = false,
            SizeToContent = SizeToContent.Manual,
            Padding = new Thickness(20)
        };

        okButton.Click += (_, _) => dialog.Close();
        await dialog.ShowDialog(this);
    }

    private static string? TryExtractPathFromMessage(string message)
    {
        if (File.Exists(message) || Directory.Exists(message))
            return message;

        int colonIndex = message.IndexOf(": ", StringComparison.Ordinal);
        if (colonIndex >= 0)
        {
            string afterColon = message[(colonIndex + 2)..].Trim();
            if (File.Exists(afterColon) || Directory.Exists(afterColon))
                return afterColon;
        }

        Match windowsMatch = WindowsPathRegex.Match(message);
        if (windowsMatch.Success)
            return windowsMatch.Groups[1].Value.Trim();

        Match uncMatch = UncPathRegex.Match(message);
        if (uncMatch.Success)
            return uncMatch.Groups[1].Value.Trim();

        return null;
    }
}
