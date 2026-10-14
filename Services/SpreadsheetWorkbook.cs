using ClosedXML.Excel;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace SEPA_Batch_Generator.Services;

internal sealed class SpreadsheetWorkbook : IDisposable
{
    private static readonly XNamespace TableNamespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    private static readonly XNamespace TextNamespace = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    private static readonly XNamespace OfficeNamespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

    private readonly XLWorkbook? _excelWorkbook;
    private readonly List<SpreadsheetWorksheet> _worksheets;

    private SpreadsheetWorkbook(XLWorkbook workbook)
    {
        _excelWorkbook = workbook;
        _worksheets = [.. workbook.Worksheets.Select(sheet => new SpreadsheetWorksheet(sheet))];
    }

    private SpreadsheetWorkbook(List<SpreadsheetWorksheet> worksheets) => _worksheets = worksheets;

    public IReadOnlyList<SpreadsheetWorksheet> Worksheets => _worksheets;

    public static bool TryOpen(string path, out SpreadsheetWorkbook? workbook, out string? errorMessage)
    {
        workbook = null;
        errorMessage = null;

        bool isOds = Path.GetExtension(path).Equals(".ods", StringComparison.OrdinalIgnoreCase);
        if (isOds)
            return TryOpenOds(path, out workbook, out errorMessage);

        if (!ExcelWorkbookLoader.TryOpenWorkbook(path, out XLWorkbook? excelWorkbook, out errorMessage))
            return false;

        workbook = new SpreadsheetWorkbook(excelWorkbook!);
        return true;
    }

    private static bool TryOpenOds(string path, out SpreadsheetWorkbook? workbook, out string? errorMessage)
    {
        workbook = null;
        errorMessage = null;

        try
        {
            using ZipArchive archive = ZipFile.OpenRead(path);
            ZipArchiveEntry contentEntry = archive.GetEntry("content.xml")
                ?? throw new InvalidDataException(Localization.Get("OdsContentMissing"));
            using Stream contentStream = contentEntry.Open();
            XDocument document = XDocument.Load(contentStream);
            List<SpreadsheetWorksheet> worksheets = [.. document
                .Descendants(TableNamespace + "table")
                .Select(ReadOdsWorksheet)];

            workbook = new SpreadsheetWorkbook(worksheets);
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = Localization.Get("OdsReadFailed", Path.GetFileName(path), ex.Message);
            return false;
        }
    }

    private static SpreadsheetWorksheet ReadOdsWorksheet(XElement tableElement)
    {
        string name = (string?)tableElement.Attribute(TableNamespace + "name") ?? "Sheet";
        Dictionary<int, List<SpreadsheetCell>> rows = [];
        int rowNumber = 1;

        foreach (XElement rowElement in tableElement.Elements(TableNamespace + "table-row"))
        {
            int repeatedRows = GetPositiveIntOrDefault(rowElement.Attribute(TableNamespace + "number-rows-repeated"));
            List<SpreadsheetCell> row = ReadOdsRow(rowElement);

            if (row.Any(cell => !cell.IsEmpty))
            {
                for (int repeat = 0; repeat < repeatedRows; repeat++)
                    rows[rowNumber + repeat] = [.. row];
            }

            rowNumber += repeatedRows;
        }

        return new SpreadsheetWorksheet(name, rows);
    }

    private static List<SpreadsheetCell> ReadOdsRow(XElement rowElement)
    {
        List<SpreadsheetCell> row = [];
        List<XElement> cellElements = [.. rowElement.Elements(TableNamespace + "table-cell")];
        int columnNumber = 1;

        for (int cellIndex = 0; cellIndex < cellElements.Count; cellIndex++)
        {
            XElement cellElement = cellElements[cellIndex];
            SpreadsheetCell cell = ReadOdsCell(cellElement);
            int repeatedColumns = GetPositiveIntOrDefault(cellElement.Attribute(TableNamespace + "number-columns-repeated"));
            bool hasLaterContent = cellElements.Skip(cellIndex + 1).Any(HasOdsContent);
            if (cell.IsEmpty && !hasLaterContent)
                break;

            while (row.Count < columnNumber - 1)
                row.Add(SpreadsheetCell.Empty);

            for (int column = 0; column < repeatedColumns; column++)
                row.Add(cell);

            columnNumber += repeatedColumns;
        }

        return row;
    }

    private static SpreadsheetCell ReadOdsCell(XElement cell)
    {
        string valueType = (string?)cell.Attribute(OfficeNamespace + "value-type") ?? string.Empty;
        string value = (string?)cell.Attribute(OfficeNamespace + "value") ?? string.Empty;
        string textValue = string.Join(string.Empty, cell.Descendants(TextNamespace + "p").Select(p => p.Value));
        bool isNumber = IsNumericOdsValueType(valueType);

        if (valueType == "date" && DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
            textValue = date.ToString(CultureInfo.CurrentCulture);
        else if (isNumber && valueType is "float" or "currency" or "percentage")
            textValue = value;
        else if (isNumber && string.IsNullOrWhiteSpace(textValue))
            textValue = value;

        return new SpreadsheetCell(textValue, isNumber, valueType == "currency" ? value : string.Empty);
    }

    private static bool IsNumericOdsValueType(string valueType)
        => valueType is "float" or "currency" or "percentage" or "date" or "time";

    private static bool HasOdsContent(XElement cell)
        => !string.IsNullOrWhiteSpace((string?)cell.Attribute(OfficeNamespace + "value"))
           || cell.Descendants(TextNamespace + "p").Any(p => !string.IsNullOrWhiteSpace(p.Value));

    private static int GetPositiveIntOrDefault(XAttribute? attribute)
        => int.TryParse(attribute?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) && value > 0 ? value : 1;

    public void Dispose() => _excelWorkbook?.Dispose();
}

internal sealed class SpreadsheetWorksheet
{
    private readonly IXLWorksheet? _excelWorksheet;
    private readonly Dictionary<int, List<SpreadsheetCell>>? _rows;

    public SpreadsheetWorksheet(IXLWorksheet worksheet) => (_excelWorksheet, Name) = (worksheet, worksheet.Name);
    public SpreadsheetWorksheet(string name, Dictionary<int, List<SpreadsheetCell>> rows) => (Name, _rows) = (name, rows);

    public string Name { get; }
    public int LastUsedRow => _excelWorksheet?.LastRowUsed()?.RowNumber()
        ?? (_rows!.Count == 0 ? 0 : _rows.Keys.Max());
    public int FirstUsedColumn => _excelWorksheet?.FirstCellUsed()?.Address.ColumnNumber
        ?? (_rows!.Count == 0 ? 1 : _rows.Values.SelectMany(row => row.Select((cell, index) => (cell, index))).Where(item => !item.cell.IsEmpty).Min(item => item.index + 1));
    public int LastUsedColumn => _excelWorksheet?.LastCellUsed()?.Address.ColumnNumber
        ?? (_rows!.Count == 0 ? 0 : _rows.Values.Max(row => row.Count));

    public SpreadsheetCell GetCell(int row, int column)
    {
        if (_excelWorksheet != null)
        {
            IXLCell cell = _excelWorksheet.Cell(row, column);
            return new SpreadsheetCell(cell.GetString(), cell.DataType == XLDataType.Number, cell.Style.NumberFormat.Format);
        }

        return row > 0 && column > 0 && _rows!.TryGetValue(row, out List<SpreadsheetCell>? cells) && column <= cells.Count
            ? cells[column - 1]
            : SpreadsheetCell.Empty;
    }
}

internal readonly record struct SpreadsheetCell(string Text, bool IsNumber, string NumberFormat)
{
    public static SpreadsheetCell Empty => new(string.Empty, false, string.Empty);
    public bool IsEmpty => string.IsNullOrWhiteSpace(Text);
}
