using System.Text;

namespace SEPA_Batch_Generator.Services;

public sealed class IniSettingsService
{
    public static Dictionary<string, string> Load(string path, string sectionName)
    {
        Dictionary<string, string> values = new(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(path))
            return values;

        string activeSection = string.Empty;
        string targetSection = NormalizeSectionName(sectionName);

        foreach (string line in File.ReadAllLines(path))
        {
            string trimmed = line.Trim();
            if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith('#') || trimmed.StartsWith(';'))
                continue;

            if (trimmed.StartsWith('[') && trimmed.EndsWith(']'))
            {
                activeSection = NormalizeSectionName(trimmed[1..^1]);
                continue;
            }

            if (!string.Equals(activeSection, targetSection, StringComparison.OrdinalIgnoreCase))
                continue;

            int index = trimmed.IndexOf('=');
            if (index <= 0)
                continue;

            string key = trimmed[..index].Trim();
            string value = trimmed[(index + 1)..].Trim();
            values[key] = value;
        }

        return values;
    }

    public static void Save(string path, string sectionName, Dictionary<string, string> values)
    {
        List<string> lines= [$"[{NormalizeSectionName(sectionName)}]"];

        foreach (var pair in values.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
            lines.Add($"{pair.Key}={pair.Value}");

        string? directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            Directory.CreateDirectory(directory);

        File.WriteAllText(path, string.Join(Environment.NewLine, lines), Encoding.UTF8);
    }

    private static string NormalizeSectionName(string sectionName)
        => string.IsNullOrWhiteSpace(sectionName) ? string.Empty : sectionName.Trim();
}
