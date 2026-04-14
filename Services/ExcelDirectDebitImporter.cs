using ClosedXML.Excel;
using SEPA_Batch_Generator.Models;
using System.Globalization;

namespace SEPA_Batch_Generator.Services
{
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
            var records = new List<DirectDebitRecord>();
            if (!File.Exists(excelPath))
            {
                messages.Add($"Excel bestand niet gevonden: {excelPath}");
                return records;
            }

            using var stream = OpenSharedReadStream(excelPath);
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (worksheet is null)
            {
                messages.Add($"Werkblad niet gevonden: {sheetName}");
                return records;
            }

            var lastUsedRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
            if (lastUsedRow <= headerRows)
            {
                messages.Add("Geen dataregels gevonden in het werkblad.");
                return records;
            }

            var filterColumnIndex = ToColumnIndexOrZero(filterColumn);

            for (var row = headerRows + 1; row <= lastUsedRow; row++)
            {
                if (filterColumnIndex > 0 && !string.IsNullOrWhiteSpace(filterValue))
                {
                    var currentFilterValue = worksheet.Cell(row, filterColumnIndex).GetString();
                    if (!string.Equals(currentFilterValue?.Trim(), filterValue.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                }

                var firstNamePart = GetCell(worksheet, row, layout.DebtorNameColumn);
                var lastNamePart = GetCell(worksheet, row, layout.DebtorLastNameColumn);
                var debtorName = BuildDebtorName(firstNamePart, lastNamePart);
                var debtorIban = GetCell(worksheet, row, layout.DebtorIbanColumn);
                var amountText = GetCell(worksheet, row, layout.AmountColumn);

                if (string.IsNullOrWhiteSpace(debtorName) && string.IsNullOrWhiteSpace(debtorIban) && string.IsNullOrWhiteSpace(amountText))
                {
                    continue;
                }

                if (!TryParseAmount(amountText, out var amount))
                {
                    messages.Add($"Rij {row}: bedrag niet leesbaar ({amountText}).");
                    amount = 0;
                }

                if (!TryParseDate(GetCell(worksheet, row, layout.MandateDateColumn), out var mandateDate))
                {
                    messages.Add($"Rij {row}: mandaatdatum niet leesbaar.");
                }

                var collectionDate = defaultCollectionDate ?? default;
                if (!defaultCollectionDate.HasValue && !TryParseDate(GetCell(worksheet, row, layout.CollectionDateColumn), out collectionDate))
                {
                    messages.Add($"Rij {row}: incassodatum niet leesbaar.");
                }

                var sequenceType = GetCell(worksheet, row, layout.SequenceTypeColumn).ToUpperInvariant();
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

        private static FileStream OpenSharedReadStream(string excelPath)
        {
            return new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        }

        private static string BuildDebtorName(string firstNamePart, string lastNamePart)
        {
            var first = ReplaceSpecialCharacters(firstNamePart.Trim());
            var last = ReplaceSpecialCharacters(lastNamePart.Trim());

            if (string.IsNullOrWhiteSpace(last))
            {
                return first;
            }

            if (string.IsNullOrWhiteSpace(first))
            {
                return last;
            }

            return $"{first} {last}";
        }

        private static string ReplaceSpecialCharacters(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return input;
            }

            return input
                .Replace("ä", "a")
                .Replace("Ä", "A")
                .Replace("ï", "i")
                .Replace("Ï", "I")
                .Replace("ë", "e")
                .Replace("Ë", "E")
                .Replace("ö", "o")
                .Replace("Ö", "O")
                .Replace("ü", "u")
                .Replace("Ü", "U");
        }

        private static string GetCell(IXLWorksheet worksheet, int row, string? column)
        {
            var index = ToColumnIndexOrZero(column);
            if (index <= 0)
            {
                return string.Empty;
            }

            return worksheet.Cell(row, index).GetString().Trim();
        }

        private static bool TryParseDate(string input, out DateTime value)
        {
            if (DateTime.TryParse(input, new CultureInfo("nl-NL"), DateTimeStyles.None, out value))
            {
                return true;
            }

            return DateTime.TryParse(input, CultureInfo.InvariantCulture, DateTimeStyles.None, out value);
        }

        private static bool TryParseAmount(string input, out decimal amount)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                amount = 0;
                return false;
            }

            // Normalize the input to handle different decimal separators
            var normalized = NormalizeDecimalSeparator(input.Trim());
            
            // Try parsing with invariant culture (uses period as decimal separator)
            return decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out amount);
        }

        private static string NormalizeDecimalSeparator(string input)
        {
            // Remove leading/trailing whitespace
            input = input.Trim();
            
            // Find the last occurrence of period or comma
            int lastPeriodIndex = input.LastIndexOf('.');
            int lastCommaIndex = input.LastIndexOf(',');

            if (lastPeriodIndex < 0 && lastCommaIndex < 0)
            {
                // No separators found, return as-is
                return input;
            }

            if (lastPeriodIndex > lastCommaIndex)
            {
                // Period is the last separator - could be thousand or decimal
                // If there are 3 or fewer digits after period, it's likely decimal
                int digitsAfterPeriod = input.Length - lastPeriodIndex - 1;
                if (digitsAfterPeriod <= 3)
                {
                    // Period is decimal separator
                    // Remove all commas (thousand separators in US format)
                    return input.Replace(",", "");
                }
                else
                {
                    // Period is thousand separator, comma is decimal
                    return input.Replace(".", "").Replace(",", ".");
                }
            }
            else
            {
                // Comma is the last separator
                // If there are 3 or fewer digits after comma, it's likely decimal
                int digitsAfterComma = input.Length - lastCommaIndex - 1;
                if (digitsAfterComma <= 3)
                {
                    // Comma is decimal separator (EU format)
                    // Remove period (thousand separator) and replace comma with period
                    return input.Replace(".", "").Replace(",", ".");
                }
                else
                {
                    // Comma is thousand separator, period is decimal (shouldn't happen if period not found)
                    return input.Replace(",", "");
                }
            }
        }

        private static int ToColumnIndexOrZero(string? column)
        {
            if (string.IsNullOrWhiteSpace(column))
            {
                return 0;
            }

            var normalized = column.Trim().ToUpperInvariant();
            var firstPart = normalized.Split([' ', '-', '|'], StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? normalized;

            if (int.TryParse(firstPart, out var numeric))
            {
                return numeric;
            }

            var sum = 0;
            foreach (var c in firstPart)
            {
                if (c < 'A' || c > 'Z')
                {
                    return 0;
                }

                sum = (sum * 26) + (c - 'A' + 1);
            }

            return sum;
        }
    }
}
