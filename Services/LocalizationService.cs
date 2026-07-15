using System.ComponentModel;
using System.Globalization;
using System.Resources;

namespace SEPA_Batch_Generator.Services;

public sealed class LocalizationService : INotifyPropertyChanged
{
    private static readonly Lazy<LocalizationService> _lazy = new(() => new LocalizationService());
    public static LocalizationService Instance => _lazy.Value;

    private readonly ResourceManager _rm = new("SEPA_Batch_Generator.Resources.Strings", typeof(LocalizationService).Assembly);
    private CultureInfo _culture = CultureInfo.CurrentUICulture;

    private LocalizationService() { }

    public event PropertyChangedEventHandler? PropertyChanged;

    public string this[string key]
    {
        get
        {
            string? value = _rm.GetString(key, _culture);
            if (!string.IsNullOrEmpty(value))
                return value;

            value = _rm.GetString(key, CultureInfo.InvariantCulture);
            return value ?? key;
        }
    }

    public CultureInfo Culture
    {
        get => _culture;
        set
        {
            if (value == null) return;
            if (Equals(value, _culture)) return;
            _culture = value;
            // Notify bindings that indexer values may have changed
            OnPropertyChanged(string.Empty);
            OnPropertyChanged("Item[]");
        }
    }

    public void SetCulture(CultureInfo culture)
    {
        Culture = culture;
    }

    public string Get(string key) => this[key];

    private void OnPropertyChanged(string name)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
