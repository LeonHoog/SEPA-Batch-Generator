using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using SEPA_Batch_Generator.ViewModels;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace SEPA_Batch_Generator.Views
{
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
            if (_viewModel is not null)
            {
                _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
            }

            _viewModel = DataContext as MainWindowViewModel;
            if (_viewModel is not null)
            {
                _viewModel.PropertyChanged += OnViewModelPropertyChanged;
            }
        }

        private async void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (_viewModel is null || e.PropertyName != nameof(MainWindowViewModel.PendingWarningMessage))
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(_viewModel.PendingWarningMessage))
            {
                return;
            }

            var message = _viewModel.PendingWarningMessage;
            _viewModel.PendingWarningMessage = string.Empty;
            await ShowWarningDialogAsync("Waarschuwing", message);
        }

        private async void BrowseExcelFile_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            var selected = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Selecteer Excel bestand",
                AllowMultiple = false,
                FileTypeFilter = new List<FilePickerFileType>
                {
                    new("Excel bestanden")
                    {
                        Patterns = new[] { "*.xlsx", "*.xlsm", "*.xls" }
                    }
                }
            });

            var file = selected.Count > 0 ? selected[0] : null;
            var localPath = file?.TryGetLocalPath();

            if (!string.IsNullOrWhiteSpace(localPath) && DataContext is MainWindowViewModel vm)
            {
                vm.ExcelPath = localPath;
            }
        }

        private async void BrowseOutputFolder_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            var folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "Selecteer output map",
                AllowMultiple = false
            });

            var folder = folders.Count > 0 ? folders[0] : null;
            var localPath = folder?.TryGetLocalPath();

            if (!string.IsNullOrWhiteSpace(localPath) && DataContext is MainWindowViewModel vm)
            {
                vm.OutputFolder = localPath;
            }
        }

        private async void BrowseLogFile_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (DataContext is not MainWindowViewModel vm)
            {
                return;
            }

            var suggestedName = string.IsNullOrWhiteSpace(vm.LogFilePath)
                ? "sepa-log.txt"
                : Path.GetFileName(vm.LogFilePath);

            var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Kies logbestand",
                SuggestedFileName = string.IsNullOrWhiteSpace(suggestedName) ? "sepa-log.txt" : suggestedName,
                DefaultExtension = "txt",
                FileTypeChoices = new List<FilePickerFileType>
                {
                    new("Tekstbestand")
                    {
                        Patterns = new[] { "*.txt" },
                        MimeTypes = new[] { "text/plain" }
                    }
                }
            });

            var localPath = file?.TryGetLocalPath();
            if (!string.IsNullOrWhiteSpace(localPath))
            {
                vm.LogFilePath = localPath;
            }
        }

        private async System.Threading.Tasks.Task ShowWarningDialogAsync(string title, string message)
        {
            var okButton = new Button
            {
                Content = "OK",
                HorizontalAlignment = HorizontalAlignment.Center,
                MinWidth = 90
            };

            var panel = new StackPanel
            {
                Margin = new Avalonia.Thickness(16),
                Spacing = 12,
                Children =
                {
                    new TextBlock { Text = message, TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                    okButton
                }
            };

            var dialog = new Window
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
            {
                return;
            }

            var message = listBox.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(message) && e.Source is Control control)
            {
                message = control.DataContext as string;
            }

            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            // Check if this is the "Totaalbedrag" message
            if (message.Contains("Totaalbedrag") && _viewModel is not null)
            {
                var breakdown = _viewModel.GetAmountBreakdown();
                await ShowBreakdownDialog(breakdown);
                return;
            }

            var path = TryExtractPathFromMessage(message);
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            if (File.Exists(path) || Directory.Exists(path))
            {
                await Clipboard!.SetTextAsync(path);
            }
        }

        private async Task ShowBreakdownDialog(string breakdown)
        {
            var textBlock = new TextBlock
            {
                Text = breakdown,
                FontFamily = new("Courier New"),
                FontSize = 12,
                Foreground = Foreground,
                Margin = new Thickness(0, 0, 0, 0)
            };

            var scrollViewer = new ScrollViewer
            {
                Content = textBlock,
                Height = 400
            };

            var okButton = new Button
            {
                Content = "OK",
                HorizontalAlignment = HorizontalAlignment.Right,
                MinWidth = 100
            };

            var panel = new StackPanel
            {
                Spacing = 12,
                Children = { scrollViewer, okButton }
            };

            var dialog = new Window
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
            {
                return message;
            }

            var colonIndex = message.IndexOf(": ", StringComparison.Ordinal);
            if (colonIndex >= 0)
            {
                var afterColon = message[(colonIndex + 2)..].Trim();
                if (File.Exists(afterColon) || Directory.Exists(afterColon))
                {
                    return afterColon;
                }
            }

            var windowsMatch = WindowsPathRegex.Match(message);
            if (windowsMatch.Success)
            {
                return windowsMatch.Groups[1].Value.Trim();
            }

            var uncMatch = UncPathRegex.Match(message);
            if (uncMatch.Success)
            {
                return uncMatch.Groups[1].Value.Trim();
            }

            return null;
        }
    }
}