using ClosedXML.Excel;
using SEPA_Batch_Generator.Models;

namespace SEPA_Batch_Generator.Services;

public sealed class ExcelMetadataService
{
    public static (List<string> WorksheetNames, string? ErrorMessage) GetWorksheetNames(string excelPath)
    {
        if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            return ([], null);

        if (ExcelWorkbookLoader.TryOpenWorkbook(excelPath, out XLWorkbook? workbook, out string? errorMessage))
        {
            using XLWorkbook loadedWorkbook = workbook ?? throw new InvalidOperationException("Workbook load succeeded without a workbook instance.");
            return ([.. loadedWorkbook.Worksheets.Select(w => w.Name)], null);
        }

        return ([], errorMessage);
    }

    public static (List<ColumnOption> ColumnOptions, string? ErrorMessage) GetColumnOptions(string excelPath, string sheetName, int headerRows)
    {
        if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            return ([], null);

        if (ExcelWorkbookLoader.TryOpenWorkbook(excelPath, out XLWorkbook? workbook, out string? errorMessage))
        {
            using XLWorkbook loadedWorkbook = workbook ?? throw new InvalidOperationException("Workbook load succeeded without a workbook instance.");
            IXLWorksheet? worksheet = loadedWorkbook.Worksheets.FirstOrDefault(ws =>
                string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (worksheet is null)
                return ([], null);

            IXLCell firstUsedCell = worksheet.FirstCellUsed();
            IXLCell lastUsedCell = worksheet.LastCellUsed();
            if (firstUsedCell is null || lastUsedCell is null)
                return ([], null);

            int firstColumn = firstUsedCell.Address.ColumnNumber;
            int lastColumn = lastUsedCell.Address.ColumnNumber;
            int headerRowNumber = headerRows <= 0 ? 1 : headerRows;

            List<ColumnOption> options = [];
            for (int col = firstColumn; col <= lastColumn; col++)
            {
                string letter = TextProcessor.ToColumnLetter(col);
                string header = worksheet.Cell(headerRowNumber, col).GetString().Trim();
                options.Add(new ColumnOption(letter, string.IsNullOrWhiteSpace(header) ? letter : $"{letter} - {header}"));
            }

            return (options, null);
        }

        return ([], errorMessage);
    }

    public static bool IsFileOpenElsewhere(string excelPath)
    {
        if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            return false;

        try
        {
            using FileStream stream = new(excelPath, FileMode.Open, FileAccess.Read, FileShare.None);
            return false;
        }
        catch (IOException) { return true; }
        catch (UnauthorizedAccessException) { return true; }
    }

}
