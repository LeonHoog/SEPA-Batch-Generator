using System.Globalization;
using System.Resources;

namespace SEPA_Batch_Generator.Services;

public static class Localization
{
    private static readonly ResourceManager _rm = new("SEPA_Batch_Generator.Resources.Strings", typeof(Localization).Assembly);

    public static string Get(string key)
    {
        string? value = _rm.GetString(key, CultureInfo.CurrentUICulture);
        if (!string.IsNullOrEmpty(value))
            return value;

        // Fallback to invariant (base) resources
        value = _rm.GetString(key, CultureInfo.InvariantCulture);
        return value ?? key;
    }

    public static string Get(string key, params object[] args)
    {
        string format = Get(key);
        return string.Format(CultureInfo.CurrentCulture, format, args);
    }
}
