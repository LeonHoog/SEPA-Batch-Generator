namespace SEPA_Batch_Generator.Models;

public sealed class ColumnOption(string id, string displayText) : IEquatable<ColumnOption>
{
    public static ColumnOption Empty { get; } = new(string.Empty, string.Empty);

    public string Id { get; } = NormalizeId(id);

    public string DisplayText { get; } = displayText?.Trim() ?? string.Empty;

    public override string ToString() => DisplayText;

    public bool Equals(ColumnOption? other)
    {
        if (other is null)
            return false;

        return string.Equals(Id, other.Id, StringComparison.OrdinalIgnoreCase);
    }

    public override bool Equals(object? obj) => obj is ColumnOption other && Equals(other);

    public override int GetHashCode() => StringComparer.OrdinalIgnoreCase.GetHashCode(Id);

    private static string NormalizeId(string? value)
        => string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim().ToUpperInvariant();
}
