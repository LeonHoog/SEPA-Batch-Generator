using ClosedXML.Excel;

namespace SEPA_Batch_Generator.Services
{
    public sealed class ExcelMetadataService
    {
        public static List<string> GetWorksheetNames(string excelPath)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                return [];
            }

            using var stream = OpenSharedReadStream(excelPath);
            using var workbook = new XLWorkbook(stream);
            return workbook.Worksheets.Select(w => w.Name).ToList();
        }

        public static List<string> GetFilterColumns(string excelPath, string sheetName, int headerRows)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                return [];
            }

            using var stream = OpenSharedReadStream(excelPath);
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (worksheet is null)
            {
                return [];
            }

            var firstUsedCell = worksheet.FirstCellUsed();
            var lastUsedCell = worksheet.LastCellUsed();
            if (firstUsedCell is null || lastUsedCell is null)
            {
                return [];
            }

            var firstColumn = firstUsedCell.Address.ColumnNumber;
            var lastColumn = lastUsedCell.Address.ColumnNumber;
            var headerRowNumber = headerRows <= 0 ? 1 : headerRows;

            var options = new List<string>();
            for (var col = firstColumn; col <= lastColumn; col++)
            {
                var letter = ToColumnLetter(col);
                var header = worksheet.Cell(headerRowNumber, col).GetString().Trim();
                options.Add(string.IsNullOrWhiteSpace(header) ? letter : $"{letter} - {header}");
            }

            return options;
        }

        public static bool IsFileOpenElsewhere(string excelPath)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                return false;
            }

            try
            {
                using var stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.None);
                return false;
            }
            catch (IOException)
            {
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                return true;
            }
        }

        private static FileStream OpenSharedReadStream(string excelPath)
        {
            return new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        }

        private static string ToColumnLetter(int number)
        {
            var result = string.Empty;
            while (number > 0)
            {
                var remainder = (number - 1) % 26;
                result = (char)('A' + remainder) + result;
                number = (number - 1) / 26;
            }

            return result;
        }
    }
}
