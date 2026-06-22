using SEPA_Batch_Generator.Models;

namespace SEPA_Batch_Generator.Services;

/// <summary>
/// Manages all column mapping operations: validation, lookup, normalization.
/// Eliminates reflection and centralizes column logic.
/// </summary>
public sealed class ColumnMappingManager(List<ColumnOption> availableColumns)
{
    private readonly List<ColumnOption> _availableColumns = availableColumns ?? [];

    /// <summary>
    /// Gets the ColumnOption for a given column ID, or null if not found.
    /// </summary>
    public ColumnOption? GetColumnOption(string? columnId)
    {
        string normalizedId = TextProcessor.NormalizeColumnId(columnId);
        if (string.IsNullOrWhiteSpace(normalizedId))
            return _availableColumns.FirstOrDefault(c => string.IsNullOrWhiteSpace(c.Id));

        return _availableColumns.FirstOrDefault(c => 
            string.Equals(c.Id, normalizedId, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Validates that a required column is properly mapped.
    /// </summary>
    public bool IsRequiredColumnValid(string? columnId)
    {
        ColumnOption? option = GetColumnOption(columnId);
        return option != null && !string.IsNullOrWhiteSpace(option.Id);
    }

    /// <summary>
    /// Validates that an optional column is either empty or properly mapped.
    /// </summary>
    public bool IsOptionalColumnValid(string? columnId)
    {
        if (string.IsNullOrWhiteSpace(columnId))
            return true;

        return GetColumnOption(columnId) != null;
    }

    /// <summary>
    /// Updates available columns (called when loading new Excel).
    /// </summary>
    public void UpdateAvailableColumns(List<ColumnOption> columns)
    {
        _availableColumns.Clear();
        _availableColumns.AddRange(columns);
    }

    /// <summary>
    /// Gets all available columns.
    /// </summary>
    public IReadOnlyList<ColumnOption> GetAvailableColumns() => _availableColumns.AsReadOnly();

    /// <summary>
    /// Validates all column mappings against a set of required and optional columns.
    /// </summary>
    public List<string> ValidateAllMappings(Dictionary<string, (string ColumnId, bool Required)> mappings)
    {
        List<string> errors = [];

        foreach (var (name, (columnId, required)) in mappings)
        {
            if (required && !IsRequiredColumnValid(columnId))
                errors.Add($"{name}: Required column not properly mapped.");
            else if (!required && !IsOptionalColumnValid(columnId))
                errors.Add($"{name}: Column mapping invalid.");
        }

        return errors;
    }
}
