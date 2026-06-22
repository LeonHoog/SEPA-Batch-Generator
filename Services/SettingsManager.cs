using SEPA_Batch_Generator.Models;

namespace SEPA_Batch_Generator.Services;

/// <summary>
/// Manages UI state settings stored in settings.ini.
/// Provides type-safe API for all non-creditor configuration.
/// </summary>
public sealed class SettingsManager
{
    private readonly string _filePath;

    public string ExcelPath { get; set; } = string.Empty;
    public string SheetName { get; set; } = "Sheet1";
    public int HeaderRows { get; set; } = 1;
    public string FilterColumn { get; set; } = string.Empty;
    public string FilterValue { get; set; } = string.Empty;
    public DateTime? GeneralCollectionDate { get; set; }
    public string OutputFolder { get; set; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SEPA_Output");
    public string LogFilePath { get; set; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SEPA_Output", "sepa-log.txt");
    public string GeneralDescription { get; set; } = string.Empty;
    public int BatchNumber { get; set; } = 1;

    // Column mappings
    public Dictionary<string, string> ColumnMappings { get; } = new(StringComparer.OrdinalIgnoreCase)
    {
        ["DebtorNameColumn"] = "A",
        ["DebtorLastNameColumn"] = string.Empty,
        ["DebtorIbanColumn"] = "B",
        ["DebtorBicColumn"] = string.Empty,
        ["AmountColumn"] = "C",
        ["MandateIdColumn"] = "D",
        ["MandateDateColumn"] = "E",
        ["CollectionDateColumn"] = "F",
        ["SequenceTypeColumn"] = "G",
        ["DescriptionColumn"] = "H",
        ["Address1Column"] = string.Empty,
        ["Address2Column"] = string.Empty
    };

    public SettingsManager(string filePath)
    {
        _filePath = filePath;
    }

    public void Load()
    {
        Dictionary<string, string> values = IniSettingsService.Load(_filePath, "SEPA");

        ExcelPath = Get(values, nameof(ExcelPath), ExcelPath);
        SheetName = Get(values, nameof(SheetName), SheetName);
        HeaderRows = ParseInt(Get(values, nameof(HeaderRows), HeaderRows.ToString()), HeaderRows);
        FilterColumn = TextProcessor.NormalizeColumnId(Get(values, nameof(FilterColumn), FilterColumn));
        FilterValue = Get(values, nameof(FilterValue), FilterValue);

        string collectionDateText = Get(values, nameof(GeneralCollectionDate), string.Empty);
        if (DateTime.TryParse(collectionDateText, out DateTime parsedDate))
            GeneralCollectionDate = new DateTime(parsedDate.Date.Ticks);

        OutputFolder = Get(values, nameof(OutputFolder), OutputFolder);
        LogFilePath = Get(values, nameof(LogFilePath), LogFilePath);
        GeneralDescription = Get(values, nameof(GeneralDescription), GeneralDescription);
        BatchNumber = ParseInt(Get(values, nameof(BatchNumber), BatchNumber.ToString()), BatchNumber);

        LoadColumnMappings(values);
    }

    public void Save()
    {
        Dictionary<string, string> values = new(StringComparer.OrdinalIgnoreCase)
        {
            [nameof(ExcelPath)] = ExcelPath,
            [nameof(SheetName)] = SheetName,
            [nameof(HeaderRows)] = HeaderRows.ToString(),
            [nameof(FilterColumn)] = FilterColumn,
            [nameof(FilterValue)] = FilterValue,
            [nameof(GeneralCollectionDate)] = GeneralCollectionDate?.ToString("yyyy-MM-dd") ?? string.Empty,
            [nameof(OutputFolder)] = OutputFolder,
            [nameof(LogFilePath)] = LogFilePath,
            [nameof(GeneralDescription)] = GeneralDescription,
            [nameof(BatchNumber)] = BatchNumber.ToString()
        };

        foreach (KeyValuePair<string, string> kvp in ColumnMappings)
            values[kvp.Key] = kvp.Value;

        IniSettingsService.Save(_filePath, "SEPA", values);
    }

    public ExcelLayoutSettings BuildExcelLayoutSettings() => new()
    {
        DebtorNameColumn = GetColumn("DebtorNameColumn"),
        DebtorLastNameColumn = GetColumn("DebtorLastNameColumn"),
        DebtorIbanColumn = GetColumn("DebtorIbanColumn"),
        DebtorBicColumn = GetColumn("DebtorBicColumn"),
        AmountColumn = GetColumn("AmountColumn"),
        MandateIdColumn = GetColumn("MandateIdColumn"),
        MandateDateColumn = GetColumn("MandateDateColumn"),
        CollectionDateColumn = GetColumn("CollectionDateColumn"),
        SequenceTypeColumn = GetColumn("SequenceTypeColumn"),
        DescriptionColumn = GetColumn("DescriptionColumn"),
        Address1Column = GetColumn("Address1Column"),
        Address2Column = GetColumn("Address2Column")
    };

    private void LoadColumnMappings(Dictionary<string, string> values)
    {
        foreach (string key in ColumnMappings.Keys)
            ColumnMappings[key] = TextProcessor.NormalizeColumnId(Get(values, key, ColumnMappings[key]));
    }

    private string GetColumn(string key)
        => Get(ColumnMappings, key, string.Empty);

    private static string Get(Dictionary<string, string> values, string key, string fallback)
        => values.TryGetValue(key, out string? value) ? value : fallback;

    private static int ParseInt(string value, int fallback)
        => int.TryParse(value, out int number) ? number : fallback;
}
