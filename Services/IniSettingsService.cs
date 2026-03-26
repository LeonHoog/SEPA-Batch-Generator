using System.Text;

namespace SEPA_Batch_Generator.Services
{
    public sealed class IniSettingsService
    {
        public static Dictionary<string, string> Load(string path)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!File.Exists(path))
            {
                return values;
            }

            foreach (var line in File.ReadAllLines(path))
            {
                var trimmed = line.Trim();
                if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith('#') || trimmed.StartsWith(';') || trimmed.StartsWith('['))
                {
                    continue;
                }

                var index = trimmed.IndexOf('=');
                if (index <= 0)
                {
                    continue;
                }

                var key = trimmed[..index].Trim();
                var value = trimmed[(index + 1)..].Trim();
                values[key] = value;
            }

            return values;
        }

        public static void Save(string path, Dictionary<string, string> values)
        {
            var lines = new List<string>
            {
                "[SEPA]"
            };

            foreach (var pair in values.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
            {
                lines.Add($"{pair.Key}={pair.Value}");
            }

            File.WriteAllText(path, string.Join(Environment.NewLine, lines), Encoding.UTF8);
        }
    }
}
