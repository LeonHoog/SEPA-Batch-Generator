using SEPA_Batch_Generator.Models;
using System.Globalization;
using System.Xml.Linq;

namespace SEPA_Batch_Generator.Services;

public sealed class SepaXmlGenerator
{
    public static string Generate(
        IReadOnlyList<DirectDebitRecord> records,
        SepaGenerationSettings settings,
        string outputDirectory,
        int batchNumber)
    {
        DirectDebitRecord first = records[0];
        DateTime requestedCollectionDate = first.CollectionDate.Date;
        string sequenceType = first.SequenceType;

        Directory.CreateDirectory(outputDirectory);

        string fileName = $"Incasso_dd_{requestedCollectionDate:yyyy-MM-dd}_{sequenceType}_08_{batchNumber:D3}.xml";
        string path = Path.Combine(outputDirectory, fileName);

        XNamespace ns = "urn:iso:std:iso:20022:tech:xsd:pain.008.001.08";

        DateTime now = DateTime.UtcNow;
        string messageId = $"MSG-{now:yyyyMMddHHmmss}-{batchNumber:D3}";
        string paymentInfoId = $"PMT-{requestedCollectionDate:yyyyMMdd}-{sequenceType}-{batchNumber:D3}";

        // Calculate control sum with explicit rounding to 2 decimal places for SEPA compliance
        decimal controlSum = Math.Round(records.Sum(r => r.Amount), 2, MidpointRounding.AwayFromZero);

        XDocument document = new(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(ns + "Document",
                new XElement(ns + "CstmrDrctDbtInitn",
                    new XElement(ns + "GrpHdr",
                        new XElement(ns + "MsgId", messageId),
                        new XElement(ns + "CreDtTm", now.ToString("yyyy-MM-ddTHH:mm:ss")),
                        new XElement(ns + "NbOfTxs", records.Count),
                        new XElement(ns + "CtrlSum", controlSum.ToString("0.00", CultureInfo.InvariantCulture)),
                        new XElement(ns + "InitgPty",
                            new XElement(ns + "Nm", settings.CreditorName)
                        )
                    ),
                    new XElement(ns + "PmtInf",
                        new XElement(ns + "PmtInfId", paymentInfoId),
                        new XElement(ns + "PmtMtd", "DD"),
                        new XElement(ns + "NbOfTxs", records.Count),
                        new XElement(ns + "CtrlSum", controlSum.ToString("0.00", CultureInfo.InvariantCulture)),
                        new XElement(ns + "PmtTpInf",
                            new XElement(ns + "SvcLvl", new XElement(ns + "Cd", "SEPA")),
                            new XElement(ns + "LclInstrm", new XElement(ns + "Cd", "CORE")),
                            new XElement(ns + "SeqTp", sequenceType)
                        ),
                        new XElement(ns + "ReqdColltnDt", requestedCollectionDate.ToString("yyyy-MM-dd")),
                        new XElement(ns + "Cdtr", new XElement(ns + "Nm", settings.CreditorName)),
                        new XElement(ns + "CdtrAcct", new XElement(ns + "Id", new XElement(ns + "IBAN", settings.CreditorIban))),
                        new XElement(ns + "CdtrAgt", new XElement(ns + "FinInstnId", new XElement(ns + "BICFI", settings.CreditorBic))),
                        new XElement(ns + "ChrgBr", "SLEV"),
                        new XElement(ns + "CdtrSchmeId",
                            new XElement(ns + "Id",
                                new XElement(ns + "PrvtId",
                                    new XElement(ns + "Othr",
                                        new XElement(ns + "Id", settings.CreditorId),
                                        new XElement(ns + "SchmeNm", new XElement(ns + "Prtry", "SEPA"))
                                    )
                                )
                            )
                        ),
                        records.Select((record, index) =>
                        {
                            string endToEnd = $"E2E-{requestedCollectionDate:yyyyMMdd}-{batchNumber:D3}-{index + 1:D5}";
                            string remittance = TextProcessor.BuildDescription(settings.GeneralDescription, record.DescriptionPart);

                            return new XElement(ns + "DrctDbtTxInf",
                                new XElement(ns + "PmtId", new XElement(ns + "EndToEndId", endToEnd)),
                                new XElement(ns + "InstdAmt", new XAttribute("Ccy", "EUR"), record.Amount.ToString("0.00", CultureInfo.InvariantCulture)),
                                new XElement(ns + "DrctDbtTx",
                                    new XElement(ns + "MndtRltdInf",
                                        new XElement(ns + "MndtId", record.MandateId),
                                        new XElement(ns + "DtOfSgntr", record.MandateSignedOn.ToString("yyyy-MM-dd"))
                                    )
                                ),
                                new XElement(ns + "DbtrAgt",
                                    new XElement(ns + "FinInstnId",
                                        string.IsNullOrWhiteSpace(record.DebtorBic)
                                            ? new XElement(ns + "Othr", new XElement(ns + "Id", "NOTPROVIDED"))
                                            : new XElement(ns + "BICFI", record.DebtorBic)
                                    )
                                ),
                                new XElement(ns + "Dbtr", new XElement(ns + "Nm", record.DebtorName)),
                                new XElement(ns + "DbtrAcct", new XElement(ns + "Id", new XElement(ns + "IBAN", record.DebtorIban))),
                                new XElement(ns + "RmtInf", new XElement(ns + "Ustrd", remittance))
                            );
                        })
                    )
                )
            )
        );

        document.Save(path);
        return path;
    }
}
