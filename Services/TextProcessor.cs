using System.Globalization;
using System.Text;

namespace SEPA_Batch_Generator.Services;

public sealed class TextProcessor
{
    private static readonly HashSet<char> AllowedCharacters =
        [.. "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/-?:().,'+ "];

    /// <summary>
    /// Converts column letter(s) (e.g., "A", "AB") to 1-based index.
    /// Also handles numeric input and normalized formats like "A - Header".
    /// </summary>
    public static int ToColumnIndex(string? column)
    {
        if (string.IsNullOrWhiteSpace(column))
            return 0;

        string normalized = column.Trim().ToUpperInvariant();
        string firstPart = normalized.Split([' ', '-', '|'], StringSplitOptions.RemoveEmptyEntries)
            .FirstOrDefault() ?? normalized;

        if (int.TryParse(firstPart, out int numeric))
            return numeric;

        int sum = 0;
        foreach (char c in firstPart)
        {
            if (c < 'A' || c > 'Z')
                return 0;
            sum = (sum * 26) + (c - 'A' + 1);
        }

        return sum;
    }

    /// <summary>
    /// Converts 1-based column index to letters (e.g., 1 → "A", 27 → "AA").
    /// </summary>
    public static string ToColumnLetter(int number)
    {
        string result = string.Empty;
        while (number > 0)
        {
            int remainder = (number - 1) % 26;
            result = (char)('A' + remainder) + result;
            number = (number - 1) / 26;
        }
        return result;
    }

    /// <summary>
    /// Normalizes column IDs: extracts first segment, uppercase, trim, empty if whitespace.
    /// </summary>
    public static string NormalizeColumnId(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        string normalized = value.Trim().ToUpperInvariant();
        string? firstPart = normalized.Split([' ', '-', '|'], StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
        return string.IsNullOrWhiteSpace(firstPart) ? string.Empty : firstPart;
    }

    /// <summary>
    /// Parses decimal amount from text with flexible format support.
    /// </summary>
    public static bool TryParseAmount(string input, out decimal amount)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            amount = 0;
            return false;
        }

        string trimmed = input.Trim();

        if (decimal.TryParse(trimmed, NumberStyles.Number, new CultureInfo("nl-NL"), out amount))
            return true;

        if (decimal.TryParse(trimmed, NumberStyles.Number, CultureInfo.InvariantCulture, out amount))
            return true;

        string normalized = trimmed.Replace(".", string.Empty).Replace(",", ".");
        return decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out amount);
    }

    /// <summary>
    /// Parses date with fallback to invariant culture (Dutch then Invariant).
    /// </summary>
    public static bool TryParseDate(string input, out DateTime value)
    {
        if (DateTime.TryParse(input, new CultureInfo("nl-NL"), DateTimeStyles.None, out value))
            return true;

        return DateTime.TryParse(input, CultureInfo.InvariantCulture, DateTimeStyles.None, out value);
    }

    /// <summary>
    /// Removes diacritics and special characters from names.
    /// </summary>
    public static string RemoveDiacritics(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return input;

        string normalized = input.Trim().Normalize(NormalizationForm.FormD);
        StringBuilder builder = new (normalized.Length);

        foreach (char character in normalized)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.NonSpacingMark)
                builder.Append(character);
        }

        return builder.ToString().Normalize(NormalizationForm.FormC);
    }

    /// <summary>
    /// Combines first and last name parts intelligently.
    /// </summary>
    public static string CombineNames(string? firstName, string? lastName)
    {
        string first = RemoveDiacritics(firstName?.Trim() ?? string.Empty);
        string last = RemoveDiacritics(lastName?.Trim() ?? string.Empty);

        if (string.IsNullOrWhiteSpace(last))
            return first;
        if (string.IsNullOrWhiteSpace(first))
            return last;
        return $"{first} {last}";
    }

    /// <summary>
    /// Checks if text contains only SEPA-allowed characters.
    /// </summary>
    public static bool ContainsOnlyAllowedChars(string value) 
        => value.All(AllowedCharacters.Contains);

    /// <summary>
    /// Builds a description from general and row-specific parts.
    /// </summary>
    public static string BuildDescription(string? general, string? rowSpecific)
    {
        return string.Join(' ', new[] { general?.Trim(), rowSpecific?.Trim() }
            .Where(v => !string.IsNullOrWhiteSpace(v)));
    }
}
