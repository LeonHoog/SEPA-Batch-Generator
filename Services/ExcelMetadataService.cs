using SEPA_Batch_Generator.Models;

namespace SEPA_Batch_Generator.Services;

public sealed class SpreadsheetMetadataService
{
    public static (List<string> WorksheetNames, string? ErrorMessage) GetWorksheetNames(string spreadsheetPath)
    {
        if (string.IsNullOrWhiteSpace(spreadsheetPath) || !File.Exists(spreadsheetPath))
            return ([], null);

        if (SpreadsheetWorkbook.TryOpen(spreadsheetPath, out SpreadsheetWorkbook? workbook, out string? errorMessage))
        {
            using SpreadsheetWorkbook loadedWorkbook = workbook ?? throw new InvalidOperationException("Workbook load succeeded without a workbook instance.");
            return ([.. loadedWorkbook.Worksheets.Select(w => w.Name)], null);
        }

        return ([], errorMessage);
    }

    public static (List<ColumnOption> ColumnOptions, string? ErrorMessage) GetColumnOptions(string spreadsheetPath, string sheetName, int headerRows)
    {
        if (string.IsNullOrWhiteSpace(spreadsheetPath) || !File.Exists(spreadsheetPath))
            return ([], null);

        if (SpreadsheetWorkbook.TryOpen(spreadsheetPath, out SpreadsheetWorkbook? workbook, out string? errorMessage))
        {
            using SpreadsheetWorkbook loadedWorkbook = workbook ?? throw new InvalidOperationException("Workbook load succeeded without a workbook instance.");
            SpreadsheetWorksheet? worksheet = loadedWorkbook.Worksheets.FirstOrDefault(ws =>
                string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (worksheet is null)
                return ([], null);

            if (worksheet.LastUsedRow == 0 || worksheet.LastUsedColumn == 0)
                return ([], null);

            int firstColumn = worksheet.FirstUsedColumn;
            int lastColumn = worksheet.LastUsedColumn;
            int headerRowNumber = headerRows <= 0 ? 1 : headerRows;

            List<ColumnOption> options = [];
            for (int col = firstColumn; col <= lastColumn; col++)
            {
                string letter = TextProcessor.ToColumnLetter(col);
                string header = worksheet.GetCell(headerRowNumber, col).Text.Trim();
                options.Add(new ColumnOption(letter, string.IsNullOrWhiteSpace(header) ? letter : $"{letter} - {header}"));
            }

            return (options, null);
        }

        return ([], errorMessage);
    }

    public static bool IsFileOpenElsewhere(string spreadsheetPath)
    {
        if (string.IsNullOrWhiteSpace(spreadsheetPath) || !File.Exists(spreadsheetPath))
            return false;

        try
        {
            using FileStream stream = new(spreadsheetPath, FileMode.Open, FileAccess.Read, FileShare.None);
            return false;
        }
        catch (IOException) { return true; }
        catch (UnauthorizedAccessException) { return true; }
    }

}
