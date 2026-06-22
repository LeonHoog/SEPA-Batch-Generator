using SEPA_Batch_Generator.Models;

namespace SEPA_Batch_Generator.Services;

/// <summary>
/// Manages creditor data stored in a separate creditor.ini file.
/// </summary>
public sealed class CreditorConfigManager(string filePath)
{
    private readonly string _filePath = filePath;

    public CreditorData Load()
    {
        if (!File.Exists(_filePath))
        {
            CreditorData defaultCreditor = new();
            Save(defaultCreditor);
            return defaultCreditor;
        }

        Dictionary<string, string> values = IniSettingsService.Load(_filePath, "CREDITOR");
        return new CreditorData(
            Get(values, "Name", string.Empty),
            Get(values, "Iban", string.Empty),
            Get(values, "Bic", string.Empty),
            Get(values, "Id", string.Empty)
        );
    }

    public void Save(CreditorData creditor)
    {
        Dictionary<string, string> values = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Name"] = creditor.Name,
            ["Iban"] = creditor.Iban,
            ["Bic"] = creditor.Bic,
            ["Id"] = creditor.Id
        };

        IniSettingsService.Save(_filePath, "CREDITOR", values);
    }

    private static string Get(Dictionary<string, string> values, string key, string fallback)
        => values.TryGetValue(key, out var value) ? value : fallback;
}
