using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SEPA_Batch_Generator.Models;
using SEPA_Batch_Generator.Services;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Text;

namespace SEPA_Batch_Generator.ViewModels
{
    public partial class MainWindowViewModel : ViewModelBase
    {
        private readonly IniSettingsService _iniSettingsService = new();
        private readonly ExcelDirectDebitImporter _excelImporter = new();
        private readonly ExcelMetadataService _excelMetadataService = new();
        private readonly SepaInputValidator _validator = new();
        private readonly SepaXmlGenerator _xmlGenerator = new();
        private readonly string _settingsPath = Path.Combine(AppContext.BaseDirectory, "settings.ini");

        private List<DirectDebitRecord> _validRecords = [];
        private int _metadataLoadVersion;
        private string _lastOpenElsewhereWarningPath = string.Empty;
        private bool _isLoadingSettings;
        private int _metadataRefreshDepth;

        public ObservableCollection<string> Messages { get; } = [];
        public ObservableCollection<string> WorksheetNames { get; } = [];
        public ObservableCollection<string> FilterColumnOptions { get; } = [];
        public ObservableCollection<string> ColumnOptions { get; } = [];

        public bool HasExcelSelected => !string.IsNullOrWhiteSpace(ExcelPath) && File.Exists(ExcelPath);
        public bool ShowCollectionDateColumnMapping => !GeneralCollectionDate.HasValue;
        public bool HasGeneralCollectionDate => GeneralCollectionDate.HasValue;
        public static DateTimeOffset MinimumCollectionDate => new(DateTime.Today.AddDays(2));

        [ObservableProperty]
        private bool isLoadingExcel;

        [ObservableProperty]
        private string pendingWarningMessage = string.Empty;

        [ObservableProperty]
        private string excelPath = string.Empty;

        [ObservableProperty]
        private string sheetName = "Sheet1";

        [ObservableProperty]
        private int headerRows = 1;

        [ObservableProperty]
        private string filterColumn = string.Empty;

        [ObservableProperty]
        private string filterValue = string.Empty;

        [ObservableProperty]
        private DateTimeOffset? generalCollectionDate;

        [ObservableProperty]
        private string outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SEPA_Output");

        [ObservableProperty]
        private string logFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SEPA_Output", "sepa-log.txt");

        [ObservableProperty]
        private string generalDescription = string.Empty;

        [ObservableProperty]
        private string creditorName = string.Empty;

        [ObservableProperty]
        private string creditorIban = string.Empty;

        [ObservableProperty]
        private string creditorBic = string.Empty;

        [ObservableProperty]
        private string creditorId = string.Empty;

        [ObservableProperty]
        private string debtorNameColumn = "A";

        [ObservableProperty]
        private string debtorLastNameColumn = string.Empty;

        [ObservableProperty]
        private string debtorIbanColumn = "B";

        [ObservableProperty]
        private string debtorBicColumn = string.Empty;

        [ObservableProperty]
        private string amountColumn = "C";

        [ObservableProperty]
        private string mandateIdColumn = "D";

        [ObservableProperty]
        private string mandateDateColumn = "E";

        [ObservableProperty]
        private string collectionDateColumn = "F";

        [ObservableProperty]
        private string sequenceTypeColumn = "G";

        [ObservableProperty]
        private string descriptionColumn = "H";

        [ObservableProperty]
        private string address1Column = string.Empty;

        [ObservableProperty]
        private string address2Column = string.Empty;

        [ObservableProperty]
        private int batchNumber = 1;

        [ObservableProperty]
        private bool inspectionSucceeded;

        [ObservableProperty]
        private string status = "Vul de instellingen in en start met Inspecteer.";

        [ObservableProperty]
        private decimal totalAmount;

        public string TotalAmountDisplay => TotalAmount.ToString("C", new CultureInfo("nl-NL"));

        public MainWindowViewModel()
        {
            _isLoadingSettings = true;
            LoadSettings();
            _isLoadingSettings = false;
            _ = ReloadExcelMetadataAsync();
        }

        partial void OnInspectionSucceededChanged(bool value)
        {
            GenerateXmlCommand.NotifyCanExecuteChanged();
        }

        partial void OnTotalAmountChanged(decimal value)
        {
            OnPropertyChanged(nameof(TotalAmountDisplay));
        }

        partial void OnGeneralCollectionDateChanged(DateTimeOffset? value)
        {
            OnPropertyChanged(nameof(ShowCollectionDateColumnMapping));
            OnPropertyChanged(nameof(HasGeneralCollectionDate));
            OnPropertyChanged(nameof(IsCollectionDateColumnInvalid));
            SaveSettingsQuietly();
        }

        [RelayCommand]
        private void ClearGeneralCollectionDate()
        {
            GeneralCollectionDate = null;
        }

        partial void OnExcelPathChanged(string value)
        {
            OnPropertyChanged(nameof(HasExcelSelected));
            _ = ReloadExcelMetadataAsync();
            SaveSettingsQuietly();
        }

        partial void OnSheetNameChanged(string value)
        {
            _ = ReloadExcelMetadataAsync(loadWorksheets: false);
            SaveSettingsQuietly();
        }

        partial void OnHeaderRowsChanged(int value)
        {
            _ = ReloadExcelMetadataAsync(loadWorksheets: false);
            SaveSettingsQuietly();
        }

        partial void OnFilterColumnChanged(string value) => SaveSettingsQuietly();
        partial void OnFilterValueChanged(string value) => SaveSettingsQuietly();
        partial void OnOutputFolderChanged(string value) => SaveSettingsQuietly();
        partial void OnLogFilePathChanged(string value) => SaveSettingsQuietly();
        partial void OnGeneralDescriptionChanged(string value) => SaveSettingsQuietly();
        partial void OnCreditorNameChanged(string value) => SaveSettingsQuietly();
        partial void OnCreditorIbanChanged(string value) => SaveSettingsQuietly();
        partial void OnCreditorBicChanged(string value) => SaveSettingsQuietly();
        partial void OnCreditorIdChanged(string value) => SaveSettingsQuietly();
        partial void OnDebtorNameColumnChanged(string value)
        {
            OnPropertyChanged(nameof(IsDebtorNameColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnDebtorLastNameColumnChanged(string value)
        {
            OnPropertyChanged(nameof(IsDebtorLastNameColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnDebtorIbanColumnChanged(string value)
        {
            OnPropertyChanged(nameof(IsDebtorIbanColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnDebtorBicColumnChanged(string value)
        {
            OnPropertyChanged(nameof(IsDebtorBicColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnAmountColumnChanged(string value)
        {
            OnPropertyChanged(nameof(IsAmountColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnMandateIdColumnChanged(string value)
        {
            OnPropertyChanged(nameof(IsMandateIdColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnMandateDateColumnChanged(string value)
        {
            OnPropertyChanged(nameof(IsMandateDateColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnCollectionDateColumnChanged(string value)
        {
            if (!string.IsNullOrWhiteSpace(value) && GeneralCollectionDate.HasValue)
            {
                GeneralCollectionDate = null;
            }

            OnPropertyChanged(nameof(IsCollectionDateColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnSequenceTypeColumnChanged(string value)
        {
            OnPropertyChanged(nameof(IsSequenceTypeColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnDescriptionColumnChanged(string value)
        {
            OnPropertyChanged(nameof(IsDescriptionColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnAddress1ColumnChanged(string value)
        {
            OnPropertyChanged(nameof(IsAddress1ColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnAddress2ColumnChanged(string value)
        {
            OnPropertyChanged(nameof(IsAddress2ColumnInvalid));
            SaveSettingsQuietly();
        }

        partial void OnBatchNumberChanged(int value) => SaveSettingsQuietly();

        [RelayCommand]
        private void SaveSettings()
        {
            SaveSettingsCore(addMessage: true);
        }

        [RelayCommand]
        private void Inspect()
        {
            Messages.Clear();
            InspectionSucceeded = false;
            TotalAmount = 0m;

            if (!ValidateGeneralInputs())
            {
                Status = "Inspectie mislukt";
                return;
            }

            var importMessages = new List<string>();
            var imported = ExcelDirectDebitImporter.Import(
                ExcelPath,
                SheetName,
                HeaderRows,
                new ExcelLayoutSettings
                {
                    DebtorNameColumn = DebtorNameColumn,
                    DebtorLastNameColumn = DebtorLastNameColumn,
                    DebtorIbanColumn = DebtorIbanColumn,
                    DebtorBicColumn = DebtorBicColumn,
                    AmountColumn = AmountColumn,
                    MandateIdColumn = MandateIdColumn,
                    MandateDateColumn = MandateDateColumn,
                    CollectionDateColumn = CollectionDateColumn,
                    SequenceTypeColumn = SequenceTypeColumn,
                    DescriptionColumn = DescriptionColumn,
                    Address1Column = Address1Column,
                    Address2Column = Address2Column
                },
                FilterColumn,
                FilterValue,
                GeneralCollectionDate?.Date,
                importMessages);

            foreach (var message in importMessages)
            {
                AddMessage(message);
            }

            var validationMessages = new List<string>();
            _validRecords = SepaInputValidator.Validate(imported, GeneralDescription, validationMessages);
            TotalAmount = _validRecords.Sum(r => r.Amount);
            foreach (var message in validationMessages)
            {
                AddMessage(message);
            }

            // Detailed summary of import results
            var rejected = SepaInputValidator.GetLastRejectedRecords();
            
            AddMessage(string.Empty);
            AddMessage("=== Samenvatting Validatie ===");
            AddMessage($"Totaal geïmporteerd: {imported.Count} regels");
            AddMessage($"Geaccepteerd: {_validRecords.Count} regels");
            AddMessage($"Niet meegenomen: {rejected.Count} regels");
            
            if (rejected.Count > 0)
            {
                AddMessage(string.Empty);
                AddMessage("--- Personen NIET meegenomen in incasso batch ---");
                foreach (var (record, reason) in rejected)
                {
                    AddMessage($"  Rij {record.RowNumber}: {record.DebtorName} - Reden: {reason}");
                }
            }

            if (_validRecords.Count == 0)
            {
                AddMessage("Geen geldige regels gevonden.");
                Status = "Inspectie mislukt";
                return;
            }

            AddMessage(string.Empty);
            AddMessage($"Totaalbedrag (alleen geaccepteerde regels): {TotalAmountDisplay}");

            if (validationMessages.Count > 0)
            {
                Status = "Inspectie bevat waarschuwingen";
                InspectionSucceeded = true;
            }
            else
            {
                InspectionSucceeded = true;
                AddMessage($"Inspectie succesvol: {_validRecords.Count} regels gevalideerd.");
                Status = "Inspectie succesvol";
            }

            WriteLog();
        }

        [RelayCommand(CanExecute = nameof(CanGenerateXml))]
        private void GenerateXml()
        {
            var settings = new SepaGenerationSettings
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
                var xmlPath = SepaXmlGenerator.Generate(group.ToList(), settings, OutputFolder, BatchNumber);
                AddMessage($"XML aangemaakt: {xmlPath}");
                BatchNumber++;
            }

            SaveSettings();
            WriteLog();
            Status = "XML generatie voltooid";
        }

        private bool CanGenerateXml() => InspectionSucceeded;

        private bool ValidateGeneralInputs()
        {
            var ok = true;

            if (string.IsNullOrWhiteSpace(ExcelPath))
            {
                AddMessage("Excel pad ontbreekt.");
                ok = false;
            }

            if (!GeneralCollectionDate.HasValue && string.IsNullOrWhiteSpace(CollectionDateColumn))
            {
                AddMessage("Kies een algemene incassodatum of stel de Excel-kolom voor incassodatum in.");
                ok = false;
            }

            if (GeneralCollectionDate.HasValue && GeneralCollectionDate.Value.Date < MinimumCollectionDate.Date)
            {
                AddMessage($"De algemene incassodatum moet vanaf {MinimumCollectionDate:dd-MM-yyyy} zijn.");
                ok = false;
            }

            if (string.IsNullOrWhiteSpace(CreditorName) || string.IsNullOrWhiteSpace(CreditorIban) || string.IsNullOrWhiteSpace(CreditorBic) || string.IsNullOrWhiteSpace(CreditorId))
            {
                AddMessage("Crediteurgegevens zijn niet volledig ingevuld.");
                ok = false;
            }

            return ok;
        }

        private async Task ReloadExcelMetadataAsync(bool loadWorksheets = true)
        {
            var currentVersion = ++_metadataLoadVersion;

            if (!HasExcelSelected)
            {
                WorksheetNames.Clear();
                FilterColumnOptions.Clear();
                ColumnOptions.Clear();
                ColumnOptions.Insert(0, string.Empty);
                NotifyColumnMappingValidationChanged();
                IsLoadingExcel = false;
                _lastOpenElsewhereWarningPath = string.Empty;
                return;
            }

            _metadataRefreshDepth++;

            try
            {
                IsLoadingExcel = true;

                if (loadWorksheets
                    && !string.Equals(_lastOpenElsewhereWarningPath, ExcelPath, StringComparison.OrdinalIgnoreCase)
                    && ExcelMetadataService.IsFileOpenElsewhere(ExcelPath))
                {
                    PendingWarningMessage = "Het geselecteerde Excel-bestand lijkt al geopend in een ander programma. De gegevens worden wel gewoon ingelezen (alleen-lezen).";
                    _lastOpenElsewhereWarningPath = ExcelPath;
                }

                List<string> worksheets = [];
                List<string> columns;

                if (loadWorksheets)
                {
                    worksheets = await Task.Run(() => ExcelMetadataService.GetWorksheetNames(ExcelPath));
                }

                var selectedSheet = SheetName;
                if (loadWorksheets && worksheets.Count > 0 && !worksheets.Contains(selectedSheet))
                {
                    selectedSheet = worksheets[0];
                }

                columns = await Task.Run(() => ExcelMetadataService.GetFilterColumns(ExcelPath, selectedSheet, HeaderRows));

                if (currentVersion != _metadataLoadVersion)
                {
                    return;
                }

                if (loadWorksheets)
                {
                    WorksheetNames.Clear();
                    foreach (var worksheet in worksheets)
                    {
                        WorksheetNames.Add(worksheet);
                    }

                    if (worksheets.Count > 0 && SheetName != selectedSheet)
                    {
                        SheetName = selectedSheet;
                    }
                }

                FilterColumnOptions.Clear();
                ColumnOptions.Clear();
                foreach (var column in columns)
                {
                    FilterColumnOptions.Add(column);
                    ColumnOptions.Add(column);
                }

                ColumnOptions.Insert(0, string.Empty);
                NotifyColumnMappingValidationChanged();
            }
            catch (Exception ex)
            {
                AddMessage($"Kolommen uitlezen mislukt: {ex.Message}");
            }
            finally
            {
                _metadataRefreshDepth--;

                if (currentVersion == _metadataLoadVersion)
                {
                    IsLoadingExcel = false;
                }
            }
        }

        private void LoadSettings()
        {
            var values = IniSettingsService.Load(_settingsPath);
            ExcelPath = Get(values, nameof(ExcelPath), ExcelPath);
            SheetName = Get(values, nameof(SheetName), SheetName);
            HeaderRows = ParseInt(Get(values, nameof(HeaderRows), HeaderRows.ToString()), HeaderRows);
            FilterColumn = Get(values, nameof(FilterColumn), FilterColumn);
            FilterValue = Get(values, nameof(FilterValue), FilterValue);
            var generalCollectionDateText = Get(values, nameof(GeneralCollectionDate), string.Empty);
            if (DateTime.TryParse(generalCollectionDateText, out var parsedCollectionDate))
            {
                GeneralCollectionDate = new DateTimeOffset(parsedCollectionDate.Date);
            }
            OutputFolder = Get(values, nameof(OutputFolder), OutputFolder);
            LogFilePath = Get(values, nameof(LogFilePath), LogFilePath);
            GeneralDescription = Get(values, nameof(GeneralDescription), GeneralDescription);
            CreditorName = Get(values, nameof(CreditorName), CreditorName);
            CreditorIban = Get(values, nameof(CreditorIban), CreditorIban);
            CreditorBic = Get(values, nameof(CreditorBic), CreditorBic);
            CreditorId = Get(values, nameof(CreditorId), CreditorId);
            DebtorNameColumn = Get(values, nameof(DebtorNameColumn), DebtorNameColumn);
            DebtorLastNameColumn = Get(values, nameof(DebtorLastNameColumn), DebtorLastNameColumn);
            DebtorIbanColumn = Get(values, nameof(DebtorIbanColumn), DebtorIbanColumn);
            DebtorBicColumn = Get(values, nameof(DebtorBicColumn), DebtorBicColumn);
            AmountColumn = Get(values, nameof(AmountColumn), AmountColumn);
            MandateIdColumn = Get(values, nameof(MandateIdColumn), MandateIdColumn);
            MandateDateColumn = Get(values, nameof(MandateDateColumn), MandateDateColumn);
            CollectionDateColumn = Get(values, nameof(CollectionDateColumn), CollectionDateColumn);
            SequenceTypeColumn = Get(values, nameof(SequenceTypeColumn), SequenceTypeColumn);
            DescriptionColumn = Get(values, nameof(DescriptionColumn), DescriptionColumn);
            Address1Column = Get(values, nameof(Address1Column), Address1Column);
            Address2Column = Get(values, nameof(Address2Column), Address2Column);
            BatchNumber = ParseInt(Get(values, nameof(BatchNumber), BatchNumber.ToString()), BatchNumber);
        }

        private void SaveSettingsQuietly()
        {
            SaveSettingsCore(addMessage: false);
        }

        private void SaveSettingsCore(bool addMessage)
        {
            if (_isLoadingSettings || _metadataRefreshDepth > 0)
            {
                return;
            }

            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                [nameof(ExcelPath)] = ExcelPath,
                [nameof(SheetName)] = SheetName,
                [nameof(HeaderRows)] = HeaderRows.ToString(),
                [nameof(FilterColumn)] = FilterColumn,
                [nameof(FilterValue)] = FilterValue,
                [nameof(GeneralCollectionDate)] = GeneralCollectionDate?.Date.ToString("yyyy-MM-dd") ?? string.Empty,
                [nameof(OutputFolder)] = OutputFolder,
                [nameof(LogFilePath)] = LogFilePath,
                [nameof(GeneralDescription)] = GeneralDescription,
                [nameof(CreditorName)] = CreditorName,
                [nameof(CreditorIban)] = CreditorIban,
                [nameof(CreditorBic)] = CreditorBic,
                [nameof(CreditorId)] = CreditorId,
                [nameof(DebtorNameColumn)] = DebtorNameColumn,
                [nameof(DebtorLastNameColumn)] = DebtorLastNameColumn,
                [nameof(DebtorIbanColumn)] = DebtorIbanColumn,
                [nameof(DebtorBicColumn)] = DebtorBicColumn,
                [nameof(AmountColumn)] = AmountColumn,
                [nameof(MandateIdColumn)] = MandateIdColumn,
                [nameof(MandateDateColumn)] = MandateDateColumn,
                [nameof(CollectionDateColumn)] = CollectionDateColumn,
                [nameof(SequenceTypeColumn)] = SequenceTypeColumn,
                [nameof(DescriptionColumn)] = DescriptionColumn,
                [nameof(Address1Column)] = Address1Column,
                [nameof(Address2Column)] = Address2Column,
                [nameof(BatchNumber)] = BatchNumber.ToString()
            };

            IniSettingsService.Save(_settingsPath, values);

            if (addMessage)
            {
                AddMessage($"Instellingen opgeslagen in {_settingsPath}");
            }
        }

        private void WriteLog()
        {
            var directory = Path.GetDirectoryName(LogFilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var sb = new StringBuilder();
            sb.AppendLine($"Tijd: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            foreach (var message in Messages)
            {
                sb.AppendLine(message);
            }

            File.AppendAllText(LogFilePath, sb.ToString() + Environment.NewLine, Encoding.UTF8);
        }

        private void AddMessage(string message)
        {
            Messages.Add(message);
        }

        private static string Get(Dictionary<string, string> values, string key, string fallback)
            => values.TryGetValue(key, out var value) ? value : fallback;

        private static int ParseInt(string value, int fallback)
            => int.TryParse(value, out var number) ? number : fallback;

        private bool IsRequiredColumnInvalid(string? column)
            => string.IsNullOrWhiteSpace(column) || !ColumnOptions.Contains(column);

        private bool IsOptionalColumnInvalid(string? column)
            => !string.IsNullOrWhiteSpace(column) && !ColumnOptions.Contains(column);

        private void NotifyColumnMappingValidationChanged()
        {
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

        public bool IsDebtorNameColumnInvalid => IsRequiredColumnInvalid(DebtorNameColumn);
        public bool IsDebtorLastNameColumnInvalid => IsOptionalColumnInvalid(DebtorLastNameColumn);
        public bool IsDebtorIbanColumnInvalid => IsRequiredColumnInvalid(DebtorIbanColumn);
        public bool IsDebtorBicColumnInvalid => IsOptionalColumnInvalid(DebtorBicColumn);
        public bool IsAmountColumnInvalid => IsRequiredColumnInvalid(AmountColumn);
        public bool IsMandateIdColumnInvalid => IsRequiredColumnInvalid(MandateIdColumn);
        public bool IsMandateDateColumnInvalid => IsRequiredColumnInvalid(MandateDateColumn);
        public bool IsCollectionDateColumnInvalid => ShowCollectionDateColumnMapping && IsRequiredColumnInvalid(CollectionDateColumn);
        public bool IsSequenceTypeColumnInvalid => IsRequiredColumnInvalid(SequenceTypeColumn);
        public bool IsDescriptionColumnInvalid => IsRequiredColumnInvalid(DescriptionColumn);
        public bool IsAddress1ColumnInvalid => IsOptionalColumnInvalid(Address1Column);
        public bool IsAddress2ColumnInvalid => IsOptionalColumnInvalid(Address2Column);

        public string GetAmountBreakdown()
        {
            if (_validRecords.Count == 0)
            {
                return "Geen geldige records beschikbaar.";
            }

            var sb = new StringBuilder();
            sb.AppendLine("=== Overzicht overboeking per persoon ===");
            sb.AppendLine();

            var sorted = _validRecords.OrderByDescending(r => r.Amount).ToList();
            foreach (var record in sorted)
            {
                var amount = record.Amount.ToString("0.00", CultureInfo.GetCultureInfo("nl-NL"));
                sb.AppendLine($"{record.DebtorName,-50} € {amount,12}");
            }

            sb.AppendLine();
            sb.AppendLine(new string('-', 65));
            var total = TotalAmount.ToString("0.00", CultureInfo.GetCultureInfo("nl-NL"));
            sb.AppendLine($"{"TOTAAL",-50} € {total,12}");

            return sb.ToString();
        }
    }
}
