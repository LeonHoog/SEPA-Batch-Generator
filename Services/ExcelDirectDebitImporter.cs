using SEPA_Batch_Generator.Models;
using System.Globalization;

namespace SEPA_Batch_Generator.Services;

public sealed class SpreadsheetDirectDebitImporter
{
    private static readonly HashSet<string> NonEuroCurrencyCodes = GetNonEuroCurrencyCodes();
    private static readonly HashSet<string> NonEuroCurrencySymbols = GetNonEuroCurrencySymbols();

    public static List<DirectDebitRecord> Import(
        string spreadsheetPath,
        string sheetName,
        int headerRows,
        SpreadsheetLayoutSettings layout,
        string? filterColumn,
        string? filterValue,
        DateTime? defaultCollectionDate,
        List<string> messages)
    {
        List<DirectDebitRecord> records = [];
        if (!File.Exists(spreadsheetPath))
        {
            messages.Add(Localization.Get("SpreadsheetFileNotFound", spreadsheetPath));
            return records;
        }

        if (!SpreadsheetWorkbook.TryOpen(spreadsheetPath, out SpreadsheetWorkbook? workbook, out string? workbookLoadError))
        {
            messages.Add(workbookLoadError ?? Localization.Get("SpreadsheetReadFailed"));
            return records;
        }

        using SpreadsheetWorkbook loadedWorkbook = workbook!;
        SpreadsheetWorksheet? worksheet = loadedWorkbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        if (worksheet is null)
        {
            messages.Add(Localization.Get("WorksheetNotFound", sheetName));
            return records;
        }

        int lastUsedRow = worksheet.LastUsedRow;
        if (lastUsedRow <= headerRows)
        {
            messages.Add(Localization.Get("NoDataRowsFound"));
            return records;
        }

        int filterColumnIndex = ToColumnIndexOrZero(filterColumn);
        for (int row = headerRows + 1; row <= lastUsedRow; row++)
        {
            if (filterColumnIndex > 0 && !string.IsNullOrWhiteSpace(filterValue))
            {
                string currentFilterValue = worksheet.GetCell(row, filterColumnIndex).Text;
                if (!string.Equals(currentFilterValue?.Trim(), filterValue.Trim(), StringComparison.OrdinalIgnoreCase))
                    continue;
            }

            string firstNamePart = GetCell(worksheet, row, layout.DebtorNameColumn);
            string lastNamePart = GetCell(worksheet, row, layout.DebtorLastNameColumn);
            string debtorName = BuildDebtorName(firstNamePart, lastNamePart);
            string debtorIban = GetCell(worksheet, row, layout.DebtorIbanColumn);
            // Let ClosedXML handle numeric cell values (locale-aware). Fallback to text parsing.
            int amountColumnIndex = ToColumnIndexOrZero(layout.AmountColumn);
            decimal amount = 0;
            string amountText = string.Empty;
            bool amountHasNonEuroCurrency = false;
            if (amountColumnIndex > 0)
            {
                SpreadsheetCell amountCell = worksheet.GetCell(row, amountColumnIndex);
                amountHasNonEuroCurrency = !string.IsNullOrWhiteSpace(amountCell.Text) && HasNonEuroCurrencyNotation(amountCell);
                if (amountCell.IsNumber && TryParseNumericAmount(amountCell.Text, out decimal numericAmount))
                    amount = numericAmount;
                else
                    amountText = amountCell.Text;
            }

            if (string.IsNullOrWhiteSpace(debtorName) && string.IsNullOrWhiteSpace(debtorIban) && string.IsNullOrWhiteSpace(amountText))
                continue;

            if (amountHasNonEuroCurrency)
                messages.Add(Localization.Get("AmountNonEuroCurrencyWarning", row));

            if (amount == 0 && !string.IsNullOrWhiteSpace(amountText))
            {
                if (!TryParseAmount(amountText, out amount))
                {
                    messages.Add(Localization.Get("AmountUnreadable", row, amountText));
                    amount = 0;
                }
            }

            // Round amount to exactly 2 decimal places for SEPA compliance
            amount = Math.Round(amount, 2, MidpointRounding.AwayFromZero);

            if (!TryParseDate(GetCell(worksheet, row, layout.MandateDateColumn), out DateTime mandateDate))
            {
                messages.Add(Localization.Get("MandateDateUnreadable", row));
            }

            DateTime collectionDate = defaultCollectionDate ?? default;
            if (!defaultCollectionDate.HasValue && !TryParseDate(GetCell(worksheet, row, layout.CollectionDateColumn), out collectionDate))
            {
                messages.Add(Localization.Get("CollectionDateUnreadable", row));
            }

            string sequenceType = GetCell(worksheet, row, layout.SequenceTypeColumn).ToUpperInvariant();
            if (string.IsNullOrWhiteSpace(sequenceType))
                sequenceType = "RCUR";

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

    private static string GetCell(SpreadsheetWorksheet worksheet, int row, string? column)
    {
        int index = ToColumnIndexOrZero(column);
        if (index <= 0)
            return string.Empty;

        return worksheet.GetCell(row, index).Text.Trim();
    }

    private static bool TryParseDate(string input, out DateTime value) 
        => TextProcessor.TryParseDate(input, out value);

    private static bool TryParseAmount(string input, out decimal amount) 
        => TextProcessor.TryParseAmount(input, out amount);

    private static bool TryParseNumericAmount(string input, out decimal amount)
        => decimal.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out amount)
           || decimal.TryParse(input, NumberStyles.Any, CultureInfo.CurrentCulture, out amount);

    private static bool HasNonEuroCurrencyNotation(SpreadsheetCell cell)
    {
        return ContainsCurrencyMarker(cell.Text) || ContainsCurrencyMarker(cell.NumberFormat);
    }

    private static bool ContainsCurrencyMarker(string value)
    {
        if (value.Contains("EUR", StringComparison.OrdinalIgnoreCase)
            || value.Contains('€', StringComparison.Ordinal))
        {
            return false;
        }

        return NonEuroCurrencyCodes.Any(code => value.Contains(code, StringComparison.OrdinalIgnoreCase))
            || NonEuroCurrencySymbols.Any(symbol => value.Contains(symbol, StringComparison.Ordinal));
    }

    private static HashSet<string> GetNonEuroCurrencyCodes()
        => CultureInfo.GetCultures(CultureTypes.SpecificCultures)
            .Select(culture => new RegionInfo(culture.Name).ISOCurrencySymbol)
            .Where(code => !string.Equals(code, "EUR", StringComparison.OrdinalIgnoreCase))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

    private static HashSet<string> GetNonEuroCurrencySymbols()
        => CultureInfo.GetCultures(CultureTypes.SpecificCultures)
            .Select(culture => new RegionInfo(culture.Name))
            .Where(region => !string.Equals(region.ISOCurrencySymbol, "EUR", StringComparison.OrdinalIgnoreCase))
            .Select(region => region.CurrencySymbol)
            .Where(symbol => !string.IsNullOrWhiteSpace(symbol))
            .ToHashSet(StringComparer.Ordinal);

    private static int ToColumnIndexOrZero(string? column) 
        => TextProcessor.ToColumnIndex(column);
}
