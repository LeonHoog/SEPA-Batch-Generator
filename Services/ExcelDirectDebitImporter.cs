using ClosedXML.Excel;
using SEPA_Batch_Generator.Models;

namespace SEPA_Batch_Generator.Services;

public sealed class ExcelDirectDebitImporter
{
    public static List<DirectDebitRecord> Import(
        string excelPath,
        string sheetName,
        int headerRows,
        ExcelLayoutSettings layout,
        string? filterColumn,
        string? filterValue,
        DateTime? defaultCollectionDate,
        List<string> messages)
    {
        List<DirectDebitRecord> records = [];
        if (!File.Exists(excelPath))
        {
            messages.Add($"Excel bestand niet gevonden: {excelPath}");
            return records;
        }

        if (!ExcelWorkbookLoader.TryOpenWorkbook(excelPath, out XLWorkbook? workbook, out string? workbookLoadError))
        {
            messages.Add(workbookLoadError ?? "Excel-bestand kon niet worden gelezen.");
            return records;
        }

        using XLWorkbook loadedWorkbook = workbook!;
        IXLWorksheet? worksheet = loadedWorkbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        if (worksheet is null)
        {
            messages.Add($"Werkblad niet gevonden: {sheetName}");
            return records;
        }

        int lastUsedRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        if (lastUsedRow <= headerRows)
        {
            messages.Add("Geen dataregels gevonden in het werkblad.");
            return records;
        }

        int filterColumnIndex = ToColumnIndexOrZero(filterColumn);
        for (int row = headerRows + 1; row <= lastUsedRow; row++)
        {
            if (filterColumnIndex > 0 && !string.IsNullOrWhiteSpace(filterValue))
            {
                string currentFilterValue = worksheet.Cell(row, filterColumnIndex).GetString();
                if (!string.Equals(currentFilterValue?.Trim(), filterValue.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
            }

            string firstNamePart = GetCell(worksheet, row, layout.DebtorNameColumn);
            string lastNamePart = GetCell(worksheet, row, layout.DebtorLastNameColumn);
            string debtorName = BuildDebtorName(firstNamePart, lastNamePart);
            string debtorIban = GetCell(worksheet, row, layout.DebtorIbanColumn);
            // Let ClosedXML handle numeric cell values (locale-aware). Fallback to text parsing.
            int amountColumnIndex = ToColumnIndexOrZero(layout.AmountColumn);
            decimal amount = 0;
            string amountText = string.Empty;
            if (amountColumnIndex > 0)
            {
                IXLCell amountCell = worksheet.Cell(row, amountColumnIndex);
                if (amountCell.DataType == XLDataType.Number)
                    amount = amountCell.GetValue<decimal>();
                else
                    amountText = amountCell.GetString();
            }

            if (string.IsNullOrWhiteSpace(debtorName) && string.IsNullOrWhiteSpace(debtorIban) && string.IsNullOrWhiteSpace(amountText))
            {
                continue;
            }

            if (amount == 0 && !string.IsNullOrWhiteSpace(amountText))
            {
                if (!TryParseAmount(amountText, out amount))
                {
                    messages.Add($"Rij {row}: bedrag niet leesbaar ({amountText}).");
                    amount = 0;
                }
            }

            // Round amount to exactly 2 decimal places for SEPA compliance
            amount = Math.Round(amount, 2, MidpointRounding.AwayFromZero);

            if (!TryParseDate(GetCell(worksheet, row, layout.MandateDateColumn), out DateTime mandateDate))
            {
                messages.Add($"Rij {row}: mandaatdatum niet leesbaar.");
            }

            DateTime collectionDate = defaultCollectionDate ?? default;
            if (!defaultCollectionDate.HasValue && !TryParseDate(GetCell(worksheet, row, layout.CollectionDateColumn), out collectionDate))
            {
                messages.Add($"Rij {row}: incassodatum niet leesbaar.");
            }

            string sequenceType = GetCell(worksheet, row, layout.SequenceTypeColumn).ToUpperInvariant();
            if (string.IsNullOrWhiteSpace(sequenceType))
            {
                sequenceType = "RCUR";
            }

            records.Add(new DirectDebitRecord
            {
                RowNumber = row,
                DebtorName = debtorName,
                DebtorIban = debtorIban.Replace(" ", string.Empty).ToUpperInvariant(),
                DebtorBic = GetCell(worksheet, row, layout.DebtorBicColumn),
                Amount = amount,
                MandateId = GetCell(worksheet, row, layout.MandateIdColumn),
                MandateSignedOn = mandateDate,
                CollectionDate = collectionDate,
                SequenceType = sequenceType,
                DescriptionPart = GetCell(worksheet, row, layout.DescriptionColumn),
                AddressLine1 = GetCell(worksheet, row, layout.Address1Column),
                AddressLine2 = GetCell(worksheet, row, layout.Address2Column)
            });
        }

        return records;
    }

    private static string BuildDebtorName(string firstNamePart, string lastNamePart) 
        => TextProcessor.CombineNames(firstNamePart, lastNamePart);

    private static string ReplaceSpecialCharacters(string input) 
        => TextProcessor.RemoveDiacritics(input);

    private static string GetCell(IXLWorksheet worksheet, int row, string? column)
    {
        int index = ToColumnIndexOrZero(column);
        if (index <= 0)
            return string.Empty;

        return worksheet.Cell(row, index).GetString().Trim();
    }

    private static bool TryParseDate(string input, out DateTime value) 
        => TextProcessor.TryParseDate(input, out value);

    private static bool TryParseAmount(string input, out decimal amount) 
        => TextProcessor.TryParseAmount(input, out amount);

    private static int ToColumnIndexOrZero(string? column) 
        => TextProcessor.ToColumnIndex(column);
}
