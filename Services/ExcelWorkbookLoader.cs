using ClosedXML.Excel;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace SEPA_Batch_Generator.Services;

internal static class ExcelWorkbookLoader
{
    private static readonly XNamespace MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    public static bool TryOpenWorkbook(string excelPath, out XLWorkbook? workbook, out string? errorMessage)
    {
        workbook = null;
        errorMessage = null;

        if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            return false;

        if (TryLoadWorkbook(excelPath, out workbook, out errorMessage))
            return true;

        if (!TryCreateLibreOfficeCompatibleCopy(excelPath, out string sanitizedPath, out string? sanitizeError))
        {
            errorMessage = sanitizeError ?? errorMessage;
            return false;
        }

        try
        {
            workbook = LoadWorkbook(sanitizedPath);
            errorMessage = null;
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = BuildFriendlyLoadErrorMessage(excelPath, ex);
            return false;
        }
        finally
        {
            TryDeleteFile(sanitizedPath);
        }
    }

    private static bool TryLoadWorkbook(string excelPath, out XLWorkbook? workbook, out string? errorMessage)
    {
        workbook = null;
        errorMessage = null;

        try
        {
            workbook = LoadWorkbook(excelPath);
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = BuildFriendlyLoadErrorMessage(excelPath, ex);
            return false;
        }
    }

    private static XLWorkbook LoadWorkbook(string excelPath)
    {
        using FileStream stream = new(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return new XLWorkbook(stream);
    }

    private static bool TryCreateLibreOfficeCompatibleCopy(string excelPath, out string sanitizedPath, out string? errorMessage)
    {
        sanitizedPath = Path.Combine(Path.GetTempPath(), $"sepa-workbook-{Guid.NewGuid():N}.xlsx");
        errorMessage = null;

        try
        {
            using ZipArchive sourceArchive = ZipFile.OpenRead(excelPath);
            using FileStream targetStream = new(sanitizedPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
            using ZipArchive targetArchive = new(targetStream, ZipArchiveMode.Create, leaveOpen: false);

            foreach (ZipArchiveEntry sourceEntry in sourceArchive.Entries)
            {
                if (IsCommentRelatedEntry(sourceEntry.FullName))
                    continue;

                ZipArchiveEntry targetEntry = targetArchive.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
                using Stream sourceEntryStream = sourceEntry.Open();
                using Stream targetEntryStream = targetEntry.Open();

                if (sourceEntry.FullName == "[Content_Types].xml")
                    RewriteContentTypes(sourceEntryStream, targetEntryStream);
                else if (IsWorksheetXml(sourceEntry.FullName))
                    RewriteWorksheetXml(sourceEntryStream, targetEntryStream);
                else if (IsWorksheetRelationshipXml(sourceEntry.FullName))
                    RewriteWorksheetRelationships(sourceEntryStream, targetEntryStream);
                else
                    sourceEntryStream.CopyTo(targetEntryStream);
            }

            return true;
        }
        catch (Exception ex)
        {
            TryDeleteFile(sanitizedPath);
            errorMessage = BuildFriendlyLoadErrorMessage(excelPath, ex);
            return false;
        }
    }

    private static void RewriteWorksheetXml(Stream source, Stream target)
    {
        XDocument document = XDocument.Load(source, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        XElement? root = document.Root;
        if (root is not null)
        {
            root.Elements(MainNamespace + "legacyDrawing").Remove();
            root.Elements(MainNamespace + "legacyDrawingHF").Remove();
        }

        SaveXml(document, target);
    }

    private static void RewriteWorksheetRelationships(Stream source, Stream target)
    {
        XDocument document = XDocument.Load(source, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        XElement? root = document.Root;
        if (root is not null)
        {
            List<XElement> relationships = root.Elements(RelationshipsNamespace + "Relationship").ToList();
            foreach (XElement relationship in relationships)
            {
                string? type = relationship.Attribute("Type")?.Value;
                if (!string.IsNullOrWhiteSpace(type) &&
                    (type.Contains("/comments", StringComparison.OrdinalIgnoreCase) ||
                     type.Contains("/vmlDrawing", StringComparison.OrdinalIgnoreCase) ||
                     type.Contains("/comment", StringComparison.OrdinalIgnoreCase)))
                {
                    relationship.Remove();
                }
            }
        }

        SaveXml(document, target);
    }

    private static void RewriteContentTypes(Stream source, Stream target)
    {
        XDocument document = XDocument.Load(source, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        XElement? root = document.Root;
        if (root is not null)
        {
            List<XElement> overrides = [.. root.Elements(ContentTypesNamespace + "Override")];
            foreach (XElement overrideElement in overrides)
            {
                string? partName = overrideElement.Attribute("PartName")?.Value;
                if (!string.IsNullOrWhiteSpace(partName) && IsCommentRelatedPart(partName))
                    overrideElement.Remove();
            }
        }

        SaveXml(document, target);
    }

    private static bool IsWorksheetXml(string fullName)
        => fullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
           && fullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
           && !fullName.Contains("/_rels/", StringComparison.OrdinalIgnoreCase);

    private static bool IsWorksheetRelationshipXml(string fullName)
        => fullName.StartsWith("xl/worksheets/_rels/", StringComparison.OrdinalIgnoreCase)
           && fullName.EndsWith(".xml.rels", StringComparison.OrdinalIgnoreCase);

    private static bool IsCommentRelatedEntry(string fullName)
        => IsCommentRelatedPart("/" + fullName.Replace('\\', '/').TrimStart('/'))
           || fullName.EndsWith(".vml", StringComparison.OrdinalIgnoreCase)
           || fullName.Contains("threadedComment", StringComparison.OrdinalIgnoreCase);

    private static bool IsCommentRelatedPart(string partName)
        => partName.StartsWith("/xl/comments", StringComparison.OrdinalIgnoreCase)
           || partName.StartsWith("/xl/threadedComments", StringComparison.OrdinalIgnoreCase)
           || partName.StartsWith("/xl/drawings/vmlDrawing", StringComparison.OrdinalIgnoreCase);

    private static void SaveXml(XDocument document, Stream target)
    {
        XmlWriterSettings settings = new()
        {
            Encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            OmitXmlDeclaration = false,
            Indent = false,
            CloseOutput = false
        };

        using XmlWriter writer = XmlWriter.Create(target, settings);
        document.Save(writer);
    }

    private static string BuildFriendlyLoadErrorMessage(string excelPath, Exception ex)
    {
        string fileName = Path.GetFileName(excelPath);
        return $"Excel-bestand '{fileName}' kan niet worden gelezen. Open het bestand opnieuw en sla het eventueel opnieuw op als .xlsx. Technische details: {ex.Message}";
    }

    private static void TryDeleteFile(string filePath)
    {
        try
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
        catch {}
    }
}