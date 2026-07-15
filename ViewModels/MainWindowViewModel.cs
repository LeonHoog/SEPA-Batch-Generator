using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SEPA_Batch_Generator.Models;
using SEPA_Batch_Generator.Services;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Text;

namespace SEPA_Batch_Generator.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    private readonly SettingsManager _settings;
    private readonly CreditorConfigManager _creditorConfig;
    private readonly ColumnMappingManager _columnMgr = new([]);
    private List<DirectDebitRecord> _validRecords = [];
    private int _metadataLoadVersion;
    private string _lastOpenElsewhereWarningPath = string.Empty;
    private bool _isLoadingSettings;
    private int _metadataRefreshDepth;
    private bool _hasLoadedColumnMetadata;
    private SepaInputValidator.ValidationResult _lastValidationResult = new([], []);

    public ObservableCollection<string> Messages { get; } = [];
    public ObservableCollection<string> WorksheetNames { get; } = [];
    public ObservableCollection<ColumnOption> FilterColumnOptions { get; } = [];
    public ObservableCollection<ColumnOption> ColumnOptions { get; } = [];

    public bool HasExcelSelected => !string.IsNullOrWhiteSpace(ExcelPath) && File.Exists(ExcelPath);
    public bool CanEditExcelMapping => HasExcelSelected && !IsLoadingExcel && _hasLoadedColumnMetadata;
    public bool ShowCollectionDateColumnMapping => !GeneralCollectionDate.HasValue;
    public bool HasGeneralCollectionDate => GeneralCollectionDate.HasValue;
    public static DateTimeOffset MinimumCollectionDate => new(DateTime.Today.AddDays(2));

    [ObservableProperty] private bool isLoadingExcel;
    [ObservableProperty] private string pendingWarningMessage = string.Empty;
    [ObservableProperty] private string excelPath = string.Empty;
    [ObservableProperty] private string sheetName = "Sheet1";
    [ObservableProperty] private int headerRows = 1;
    [ObservableProperty] private string filterColumn = string.Empty;
    [ObservableProperty] private string filterValue = string.Empty;
    [ObservableProperty] private DateTimeOffset? generalCollectionDate;
    [ObservableProperty] private string outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SEPA_Output");
    [ObservableProperty] private string logFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SEPA_Output", "sepa-log.txt");
    [ObservableProperty] private string generalDescription = string.Empty;
    [ObservableProperty] private string creditorName = string.Empty;
    [ObservableProperty] private string creditorIban = string.Empty;
    [ObservableProperty] private string creditorBic = string.Empty;
    [ObservableProperty] private string creditorId = string.Empty;
    [ObservableProperty] private int batchNumber = 1;
    [ObservableProperty] private bool inspectionSucceeded;
    [ObservableProperty] private string status = Localization.Get("StatusInitial");
    [ObservableProperty] private decimal totalAmount;

    public string TotalAmountDisplay => TotalAmount.ToString("C", CultureInfo.CurrentCulture);

    public LocalizationService Loc => LocalizationService.Instance;

    public MainWindowViewModel()
    {
        string settingsPath = Path.Combine(AppContext.BaseDirectory, "settings.ini");
        string creditorPath = Path.Combine(AppContext.BaseDirectory, "creditor.ini");

        _settings = new SettingsManager(settingsPath);
        _creditorConfig = new CreditorConfigManager(creditorPath);

        _isLoadingSettings = true;
        LoadSettings();
        _isLoadingSettings = false;
        _ = ReloadExcelMetadataAsync();
    }

    partial void OnInspectionSucceededChanged(bool value) => GenerateXmlCommand.NotifyCanExecuteChanged();
    partial void OnTotalAmountChanged(decimal value) => OnPropertyChanged(nameof(TotalAmountDisplay));
    partial void OnIsLoadingExcelChanged(bool value) => OnPropertyChanged(nameof(CanEditExcelMapping));

    partial void OnGeneralCollectionDateChanged(DateTimeOffset? value)
    {
        OnPropertyChanged(nameof(ShowCollectionDateColumnMapping));
        OnPropertyChanged(nameof(HasGeneralCollectionDate));
        SaveSettingsQuietly();
    }

    partial void OnExcelPathChanged(string value)
    {
        OnPropertyChanged(nameof(HasExcelSelected));
        OnPropertyChanged(nameof(CanEditExcelMapping));
        if (_isLoadingSettings) return;
        SaveSettingsCore(addMessage: false);
        ReloadSettingsFromDisk();
        _ = ReloadExcelMetadataAsync();
    }

    partial void OnSheetNameChanged(string value)
    {
        OnPropertyChanged(nameof(CanEditExcelMapping));
        if (_isLoadingSettings) return;
        SaveSettingsCore(addMessage: false);
        ReloadSettingsFromDisk();
        // Invalidate previous inspection when the user switches sheets
        InspectionSucceeded = false;
        _ = ReloadExcelMetadataAsync(loadWorksheets: false);
    }

    partial void OnHeaderRowsChanged(int value)
    {
        OnPropertyChanged(nameof(CanEditExcelMapping));
        _ = ReloadExcelMetadataAsync(loadWorksheets: false);
        SaveSettingsQuietly();
    }

    partial void OnFilterColumnChanged(string value) => SaveSettingsQuietly();
    partial void OnFilterValueChanged(string value) => SaveSettingsQuietly();
    partial void OnOutputFolderChanged(string value) => SaveSettingsQuietly();
    partial void OnLogFilePathChanged(string value) => SaveSettingsQuietly();
    partial void OnGeneralDescriptionChanged(string value) => SaveSettingsQuietly();
    partial void OnCreditorNameChanged(string value)
    {
        SaveSettingsQuietly();
        SaveCreditorQuietly();
    }

    partial void OnCreditorIbanChanged(string value)
    {
        SaveSettingsQuietly();
        SaveCreditorQuietly();
    }

    partial void OnCreditorBicChanged(string value)
    {
        SaveSettingsQuietly();
        SaveCreditorQuietly();
    }

    partial void OnCreditorIdChanged(string value)
    {
        SaveSettingsQuietly();
        SaveCreditorQuietly();
    }
    partial void OnBatchNumberChanged(int value) => SaveSettingsQuietly();

    [RelayCommand]
    private void ClearGeneralCollectionDate() => GeneralCollectionDate = null;

    [RelayCommand]
    private void SaveSettings()
    {
        SaveSettingsCore(addMessage: true);
        _creditorConfig.Save(new CreditorData(CreditorName, CreditorIban, CreditorBic, CreditorId));
    }

    [RelayCommand]
    private void Inspect()
    {
        Messages.Clear();
        InspectionSucceeded = false;
        TotalAmount = 0m;

        if (!ValidateGeneralInputs())
        {
            Status = Localization.Get("StatusInspectFailed");
            return;
        }

        List<string> importMessages = [];
        List<DirectDebitRecord> imported = ExcelDirectDebitImporter.Import(
            ExcelPath, SheetName, HeaderRows, _settings.BuildExcelLayoutSettings(),
            FilterColumn, FilterValue, GeneralCollectionDate?.Date, importMessages);
        importMessages.ForEach(m => AddMessage(m));

        List<string> validationMessages = [];
        _lastValidationResult = SepaInputValidator.Validate(imported, GeneralDescription, validationMessages);
        _validRecords = _lastValidationResult.Valid;
        TotalAmount = _validRecords.Sum(r => r.Amount);
        validationMessages.ForEach(m => AddMessage(m));

        if (_validRecords.Count == 0)
        {
            AddMessage(Localization.Get("NoValidRecordsFound"));
            Status = Localization.Get("StatusInspectFailed");
            WriteLog();
            return;
        }

        LogImportSummary(imported, validationMessages);
        InspectionSucceeded = true;
        Status = validationMessages.Count > 0 ? Localization.Get("StatusInspectContainsWarnings") : Localization.Get("StatusInspectSuccess");
        WriteLog();
    }

    private void LogImportSummary(List<DirectDebitRecord> imported, List<string> validationMessages)
    {
        AddMessage(string.Empty);
        AddMessage(Localization.Get("LogSummaryHeader"));
        AddMessage(string.Format(Localization.Get("LogImported"), imported.Count));
        AddMessage(string.Format(Localization.Get("LogAccepted"), _validRecords.Count));
        AddMessage(string.Format(Localization.Get("LogRejected"), _lastValidationResult.Rejected.Count));

        if (_lastValidationResult.Rejected.Count > 0)
        {
            AddMessage(string.Empty);
            AddMessage(Localization.Get("LogRejectedHeader"));
            foreach ((DirectDebitRecord? record, string? reason) in _lastValidationResult.Rejected)
                AddMessage(string.Format(Localization.Get("LogRejectedRow"), record.RowNumber, record.DebtorName, reason));
        }

        AddMessage(string.Empty);
        AddMessage(string.Format(Localization.Get("LogTotalAmount"), TotalAmountDisplay));
        if (validationMessages.Count == 0)
            AddMessage(string.Format(Localization.Get("LogInspectionSuccessWithCount"), _validRecords.Count));
    }

    [RelayCommand(CanExecute = nameof(CanGenerateXml))]
    private void GenerateXml()
    {
        GenerateXmlFiles();
        SaveSettings();
        WriteLog();
        Status = Localization.Get("StatusXmlGenerationCompleted");
    }

    private void GenerateXmlFiles()
    {
        SepaGenerationSettings settings = new()
        {
            CreditorName = CreditorName.Trim(),
            CreditorIban = CreditorIban.Replace(" ", string.Empty).ToUpperInvariant(),
            CreditorBic = CreditorBic.Trim().ToUpperInvariant(),
            CreditorId = CreditorId.Trim(),
            GeneralDescription = GeneralDescription.Trim()
        };

        var groups = _validRecords
            .GroupBy(r => new { r.CollectionDate.Date, Seq = r.SequenceType })
            .OrderBy(g => g.Key.Date)
            .ThenBy(g => g.Key.Seq)
            .ToList();

        foreach (var group in groups)
        {
            string xmlPath = SepaXmlGenerator.Generate(group.ToList(), settings, OutputFolder, BatchNumber);
            AddMessage(string.Format(Localization.Get("XmlCreated"), xmlPath));
            BatchNumber++;
        }
    }

    private bool CanGenerateXml() => InspectionSucceeded;

    private bool ValidateGeneralInputs()
    {
        bool ok = true;

        if (string.IsNullOrWhiteSpace(ExcelPath))
        {
            AddMessage(Localization.Get("Validate_ExcelPathMissing"));
            return false; // No need to continue when there's no input data provided
        }

        if (!GeneralCollectionDate.HasValue && string.IsNullOrWhiteSpace(_settings.ColumnMappings["CollectionDateColumn"]))
        {
            AddMessage(Localization.Get("Validate_CollectionDateMissing"));
            ok = false;
        }

        if (!CanEditExcelMapping)
        {
            AddMessage(Localization.Get("Validate_WaitExcelLoad"));
            ok = false;
        }

        if (GeneralCollectionDate.HasValue && GeneralCollectionDate.Value.Date < MinimumCollectionDate.Date)
        {
            AddMessage(string.Format(Localization.Get("Validate_CollectionDateTooEarly"), MinimumCollectionDate.ToString("dd-MM-yyyy")));
            ok = false;
        }

        if (string.IsNullOrWhiteSpace(CreditorName) || string.IsNullOrWhiteSpace(CreditorIban) || 
            string.IsNullOrWhiteSpace(CreditorBic) || string.IsNullOrWhiteSpace(CreditorId))
        {
            AddMessage(Localization.Get("Validate_CreditorDataIncomplete"));
            ok = false;
        }

        return ok;
    }

    private async Task ReloadExcelMetadataAsync(bool loadWorksheets = true)
    {
        int currentVersion = ++_metadataLoadVersion;

        if (!HasExcelSelected)
        {
            WorksheetNames.Clear();
            FilterColumnOptions.Clear();
            ColumnOptions.Clear();
            _hasLoadedColumnMetadata = false;
            OnPropertyChanged(nameof(CanEditExcelMapping));
            IsLoadingExcel = false;
            _lastOpenElsewhereWarningPath = string.Empty;
            return;
        }

        _metadataRefreshDepth++;

        try
        {
            IsLoadingExcel = true;

            if (loadWorksheets && !string.Equals(_lastOpenElsewhereWarningPath, ExcelPath, StringComparison.OrdinalIgnoreCase)
                && ExcelMetadataService.IsFileOpenElsewhere(ExcelPath))
            {
                PendingWarningMessage = Localization.Get("ExcelOpenElsewhereWarning");
                _lastOpenElsewhereWarningPath = ExcelPath;
            }

            List<string> worksheets = [];
            if (loadWorksheets)
            {
                (List<string> WorksheetNames, string? ErrorMessage) worksheetResult = await Task.Run(() => ExcelMetadataService.GetWorksheetNames(ExcelPath));
                worksheets = worksheetResult.WorksheetNames;
                if (!string.IsNullOrWhiteSpace(worksheetResult.ErrorMessage))
                {
                    _hasLoadedColumnMetadata = false;
                    WorksheetNames.Clear();
                    FilterColumnOptions.Clear();
                    ColumnOptions.Clear();
                    OnPropertyChanged(nameof(CanEditExcelMapping));
                    AddMessage(worksheetResult.ErrorMessage);
                    return;
                }
            }

            string selectedSheet = SheetName;
            if (loadWorksheets && worksheets.Count > 0 && !worksheets.Contains(selectedSheet))
            {
                selectedSheet = worksheets[0];
            }

            (List<ColumnOption> ColumnOptions, string? ErrorMessage) columnResult = await Task.Run(() => ExcelMetadataService.GetColumnOptions(ExcelPath, selectedSheet, HeaderRows));
            List<ColumnOption> columns = columnResult.ColumnOptions;
            if (!string.IsNullOrWhiteSpace(columnResult.ErrorMessage))
            {
                _hasLoadedColumnMetadata = false;
                FilterColumnOptions.Clear();
                ColumnOptions.Clear();
                OnPropertyChanged(nameof(CanEditExcelMapping));
                AddMessage(columnResult.ErrorMessage);
                return;
            }

            if (currentVersion != _metadataLoadVersion)
                return;

            if (loadWorksheets)
            {
                WorksheetNames.Clear();
                foreach (string ws in worksheets)
                    WorksheetNames.Add(ws);

                if (worksheets.Count > 0 && SheetName != selectedSheet)
                    SheetName = selectedSheet;
            }

            FilterColumnOptions.Clear();
            ColumnOptions.Clear();
            FilterColumnOptions.Add(ColumnOption.Empty);
            ColumnOptions.Add(ColumnOption.Empty);
            foreach (ColumnOption col in columns)
            {
                FilterColumnOptions.Add(col);
                ColumnOptions.Add(col);
            }

            _columnMgr.UpdateAvailableColumns(columns);
            _hasLoadedColumnMetadata = true;
            // Notify UI that column option properties may have changed so saved mappings become visible
            OnPropertyChanged(nameof(FilterColumnOption));
            OnPropertyChanged(nameof(DebtorNameColumnOption));
            OnPropertyChanged(nameof(DebtorLastNameColumnOption));
            OnPropertyChanged(nameof(DebtorIbanColumnOption));
            OnPropertyChanged(nameof(DebtorBicColumnOption));
            OnPropertyChanged(nameof(AmountColumnOption));
            OnPropertyChanged(nameof(MandateIdColumnOption));
            OnPropertyChanged(nameof(MandateDateColumnOption));
            OnPropertyChanged(nameof(CollectionDateColumnOption));
            OnPropertyChanged(nameof(SequenceTypeColumnOption));
            OnPropertyChanged(nameof(DescriptionColumnOption));
            OnPropertyChanged(nameof(Address1ColumnOption));
            OnPropertyChanged(nameof(Address2ColumnOption));
            OnPropertyChanged(nameof(CanEditExcelMapping));

            // Also refresh validation flags so UI stops showing stale invalid state
            OnPropertyChanged(nameof(IsFilterColumnInvalid));
            OnPropertyChanged(nameof(IsDebtorNameColumnInvalid));
            OnPropertyChanged(nameof(IsDebtorLastNameColumnInvalid));
            OnPropertyChanged(nameof(IsDebtorIbanColumnInvalid));
            OnPropertyChanged(nameof(IsDebtorBicColumnInvalid));
            OnPropertyChanged(nameof(IsAmountColumnInvalid));
            OnPropertyChanged(nameof(IsMandateIdColumnInvalid));
            OnPropertyChanged(nameof(IsMandateDateColumnInvalid));
            OnPropertyChanged(nameof(IsCollectionDateColumnInvalid));
            OnPropertyChanged(nameof(IsSequenceTypeColumnInvalid));
            OnPropertyChanged(nameof(IsDescriptionColumnInvalid));
            OnPropertyChanged(nameof(IsAddress1ColumnInvalid));
            OnPropertyChanged(nameof(IsAddress2ColumnInvalid));
        }
        catch (Exception ex)
        {
            _hasLoadedColumnMetadata = false;
            FilterColumnOptions.Clear();
            ColumnOptions.Clear();
            OnPropertyChanged(nameof(CanEditExcelMapping));
            AddMessage(string.Format(Localization.Get("LoadingColumnsFailed"), ex.Message));
        }
        finally
        {
            _metadataRefreshDepth--;
            if (currentVersion == _metadataLoadVersion)
                IsLoadingExcel = false;
        }
    }

    private void LoadSettings()
    {
        _settings.Load();
        ApplySettingsToViewModel(_settings);

        CreditorData creditor = _creditorConfig.Load();
        CreditorName = creditor.Name;
        CreditorIban = creditor.Iban;
        CreditorBic = creditor.Bic;
        CreditorId = creditor.Id;
    }

    private void ReloadSettingsFromDisk()
    {
        _isLoadingSettings = true;
        try
        {
            _settings.Load();
            ApplySettingsToViewModel(_settings);
        }
        finally
        {
            _isLoadingSettings = false;
        }
    }

    private void SaveSettingsQuietly() => SaveSettingsCore(addMessage: false);

    private void SaveCreditorQuietly()
    {
        if (_isLoadingSettings)
            return;

        _creditorConfig.Save(new CreditorData(CreditorName, CreditorIban, CreditorBic, CreditorId));
    }

    private void SaveSettingsCore(bool addMessage)
    {
        if (_isLoadingSettings)
            return;

        ApplyViewModelToSettings(_settings);
        _settings.Save();

        if (addMessage)
            AddMessage(string.Format(Localization.Get("SettingsSaved"), Path.Combine(AppContext.BaseDirectory, "settings.ini")));
    }

    private void ApplySettingsToViewModel(SettingsManager settings)
    {
        ExcelPath = settings.ExcelPath;
        SheetName = settings.SheetName;
        HeaderRows = settings.HeaderRows;
        FilterColumn = TextProcessor.NormalizeColumnId(settings.FilterColumn);
        FilterValue = settings.FilterValue;
        GeneralCollectionDate = settings.GeneralCollectionDate.HasValue ? new DateTimeOffset(settings.GeneralCollectionDate.Value) : null;
        OutputFolder = settings.OutputFolder;
        LogFilePath = settings.LogFilePath;
        GeneralDescription = settings.GeneralDescription;
        BatchNumber = settings.BatchNumber;
    }

    private void ApplyViewModelToSettings(SettingsManager settings)
    {
        settings.ExcelPath = ExcelPath;
        settings.SheetName = SheetName;
        settings.HeaderRows = HeaderRows;
        settings.FilterColumn = FilterColumn;
        settings.FilterValue = FilterValue;
        settings.GeneralCollectionDate = GeneralCollectionDate?.Date;
        settings.OutputFolder = OutputFolder;
        settings.LogFilePath = LogFilePath;
        settings.GeneralDescription = GeneralDescription;
        settings.BatchNumber = BatchNumber;
    }

    private void WriteLog()
    {
        string? directory = Path.GetDirectoryName(LogFilePath);
        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);

        StringBuilder sb = new();
        sb.AppendLine(string.Format(Localization.Get("LogTimePrefix"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
        foreach (string message in Messages)
            sb.AppendLine(message);

        File.AppendAllText(LogFilePath, sb.ToString() + Environment.NewLine, Encoding.UTF8);
    }

    private void AddMessage(string message) => Messages.Add(message);

    public string GetAmountBreakdown()
    {
        if (_validRecords.Count == 0)
            return Localization.Get("NoValidRecordsAvailable");

        StringBuilder sb = new();
        sb.AppendLine(Localization.Get("AmountOverviewHeader"));
        sb.AppendLine();

        List<DirectDebitRecord> sorted = _validRecords.OrderByDescending(r => r.Amount).ToList();
        foreach (DirectDebitRecord? record in sorted)
        {
            string amount = record.Amount.ToString("0.00", CultureInfo.CurrentCulture);
                sb.AppendLine($"{record.DebtorName,-50} € {amount,12}");
        }

        sb.AppendLine();
        sb.AppendLine(new string('-', 65));
        string total = TotalAmount.ToString("0.00", CultureInfo.CurrentCulture);
        sb.AppendLine($"{Localization.Get("TOTAL_LABEL"),-50} € {total,12}");

        return sb.ToString();
    }

    // Column mapping binding properties
    public ColumnOption? FilterColumnOption 
    { 
        get => GetColumnOption(nameof(FilterColumn)); 
        set => SetColumnProperty(nameof(FilterColumn), value); 
    }

    public ColumnOption? DebtorNameColumnOption 
    { 
        get => GetColumnOption("DebtorNameColumn"); 
        set => SetColumnProperty("DebtorNameColumn", value); 
    }

    public ColumnOption? DebtorLastNameColumnOption 
    { 
        get => GetColumnOption("DebtorLastNameColumn"); 
        set => SetColumnProperty("DebtorLastNameColumn", value); 
    }

    public ColumnOption? DebtorIbanColumnOption 
    { 
        get => GetColumnOption("DebtorIbanColumn"); 
        set => SetColumnProperty("DebtorIbanColumn", value); 
    }

    public ColumnOption? DebtorBicColumnOption 
    { 
        get => GetColumnOption("DebtorBicColumn"); 
        set => SetColumnProperty("DebtorBicColumn", value); 
    }

    public ColumnOption? AmountColumnOption 
    { 
        get => GetColumnOption("AmountColumn"); 
        set => SetColumnProperty("AmountColumn", value); 
    }

    public ColumnOption? MandateIdColumnOption 
    { 
        get => GetColumnOption("MandateIdColumn"); 
        set => SetColumnProperty("MandateIdColumn", value); 
    }

    public ColumnOption? MandateDateColumnOption 
    { 
        get => GetColumnOption("MandateDateColumn"); 
        set => SetColumnProperty("MandateDateColumn", value); 
    }

    public ColumnOption? CollectionDateColumnOption 
    { 
        get => GetColumnOption("CollectionDateColumn"); 
        set => SetColumnProperty("CollectionDateColumn", value); 
    }

    public ColumnOption? SequenceTypeColumnOption 
    { 
        get => GetColumnOption("SequenceTypeColumn"); 
        set => SetColumnProperty("SequenceTypeColumn", value); 
    }

    public ColumnOption? DescriptionColumnOption 
    { 
        get => GetColumnOption("DescriptionColumn"); 
        set => SetColumnProperty("DescriptionColumn", value); 
    }

    public ColumnOption? Address1ColumnOption 
    { 
        get => GetColumnOption("Address1Column"); 
        set => SetColumnProperty("Address1Column", value); 
    }

    public ColumnOption? Address2ColumnOption 
    { 
        get => GetColumnOption("Address2Column"); 
        set => SetColumnProperty("Address2Column", value); 
    }

    // Column validation properties
    public bool IsFilterColumnInvalid => IsColumnInvalid(nameof(FilterColumn), required: false);
    public bool IsDebtorNameColumnInvalid => IsColumnInvalid("DebtorNameColumn", required: true);
    public bool IsDebtorLastNameColumnInvalid => IsColumnInvalid("DebtorLastNameColumn", required: false);
    public bool IsDebtorIbanColumnInvalid => IsColumnInvalid("DebtorIbanColumn", required: true);
    public bool IsDebtorBicColumnInvalid => IsColumnInvalid("DebtorBicColumn", required: false);
    public bool IsAmountColumnInvalid => IsColumnInvalid("AmountColumn", required: true);
    public bool IsMandateIdColumnInvalid => IsColumnInvalid("MandateIdColumn", required: true);
    public bool IsMandateDateColumnInvalid => IsColumnInvalid("MandateDateColumn", required: true);
    public bool IsCollectionDateColumnInvalid => ShowCollectionDateColumnMapping && IsColumnInvalid("CollectionDateColumn", required: true);
    public bool IsSequenceTypeColumnInvalid => IsColumnInvalid("SequenceTypeColumn", required: true);
    public bool IsDescriptionColumnInvalid => IsColumnInvalid("DescriptionColumn", required: false);
    public bool IsAddress1ColumnInvalid => IsColumnInvalid("Address1Column", required: false);
    public bool IsAddress2ColumnInvalid => IsColumnInvalid("Address2Column", required: false);

    private void SetColumnProperty(string settingsKey, ColumnOption? value)
    {
        if (value is null) return;
        _settings.ColumnMappings[settingsKey] = TextProcessor.NormalizeColumnId(value.Id);
        OnPropertyChanged(settingsKey + "Option");
        OnPropertyChanged("Is" + settingsKey + "Invalid");
        SaveSettingsQuietly();
    }

    private ColumnOption? GetColumnOption(string settingsKey)
    {
        if (settingsKey == nameof(FilterColumn))
            return _columnMgr.GetColumnOption(FilterColumn);

        return _columnMgr.GetColumnOption(_settings.ColumnMappings[settingsKey]);
    }

    private bool IsColumnInvalid(string settingsKey, bool required)
    {
        string columnId = settingsKey == nameof(FilterColumn)
            ? FilterColumn
            : _settings.ColumnMappings[settingsKey];

        return required
            ? !_columnMgr.IsRequiredColumnValid(columnId)
            : !_columnMgr.IsOptionalColumnValid(columnId);
    }
}
