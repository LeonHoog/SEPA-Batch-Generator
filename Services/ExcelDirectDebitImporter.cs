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
            var first = firstNamePart.Trim();
            var last = lastNamePart.Trim();

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
            if (decimal.TryParse(input, NumberStyles.Number, new CultureInfo("nl-NL"), out amount))
            {
                return true;
            }

            return decimal.TryParse(input, NumberStyles.Number, CultureInfo.InvariantCulture, out amount);
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
