using SEPA_Batch_Generator.Models;
using System.Text.RegularExpressions;
using IbanNet;

namespace SEPA_Batch_Generator.Services;

public sealed class SepaInputValidator
{
    public record ValidationResult(List<DirectDebitRecord> Valid, List<(DirectDebitRecord Record, string Reason)> Rejected);

    public static ValidationResult Validate(List<DirectDebitRecord> importedRecords, string generalDescription, List<string> messages)
    {
        List<DirectDebitRecord> valid = [];
        List<(DirectDebitRecord Record, string Reason)> rejected = [];

        if (string.IsNullOrWhiteSpace(generalDescription) && importedRecords.All(r => string.IsNullOrWhiteSpace(r.DescriptionPart)))
        {
            messages.Add("WAARSCHUWING: Algemene omschrijving en regel-omschrijving mogen niet beide leeg zijn.");
        }

        foreach (DirectDebitRecord record in importedRecords)
        {
            List<string> errors = ValidateRecord(record, generalDescription, messages);
            if (errors.Count > 0)
                rejected.Add((record, string.Join("; ", errors)));
            else
                valid.Add(record);
        }

        return new ValidationResult(valid, rejected);
    }

    private static List<string> ValidateRecord(DirectDebitRecord record, string generalDescription, List<string> messages)
    {
        string rowPrefix = $"Rij {record.RowNumber}:";
        List<string> errors = [];

        if (string.IsNullOrWhiteSpace(record.DebtorName))
        {
            messages.Add($"{rowPrefix} WAARSCHUWING: Naam ontbreekt.");
            errors.Add("Naam ontbreekt");
        }
        else if (!TextProcessor.ContainsOnlyAllowedChars(record.DebtorName))
        {
            messages.Add($"{rowPrefix} WAARSCHUWING: Naam bevat ongeoorloofde tekens.");
            errors.Add("Naam bevat ongeoorloofde tekens");
        }

        if (!new IbanValidator().Validate(record.DebtorIban).IsValid)
        {
            messages.Add($"{rowPrefix} WAARSCHUWING: IBAN is ongeldig ({record.DebtorIban}).");
            errors.Add("IBAN ongeldig");
        }

        if (record.Amount <= 0)
        {
            messages.Add($"{rowPrefix} WAARSCHUWING: Bedrag mag niet 0 of negatief zijn (bedrag: {record.Amount}).");
            errors.Add("Bedrag 0 of negatief");
        }

        if (record.CollectionDate.Date < DateTime.Today)
        {
            messages.Add($"{rowPrefix} WAARSCHUWING: Incassodatum ligt in het verleden ({record.CollectionDate:yyyy-MM-dd}).");
            errors.Add("Incassodatum in verleden");
        }

        bool mandateIdMissing = string.IsNullOrWhiteSpace(record.MandateId);
        bool mandateDateMissing = record.MandateSignedOn == default;

        if (mandateIdMissing || mandateDateMissing)
        {
            if (mandateIdMissing && mandateDateMissing)
            {
                messages.Add($"{rowPrefix} WAARSCHUWING: Zowel machtigingskenmerk als ondertekeningsdatum ontbreken.");
                errors.Add("Machtigingskenmerk en -datum ontbreken");
            }
            else if (mandateIdMissing)
            {
                messages.Add($"{rowPrefix} WAARSCHUWING: Machtigingskenmerk ontbreekt.");
                errors.Add("Machtigingskenmerk ontbreekt");
            }
            else
            {
                messages.Add($"{rowPrefix} WAARSCHUWING: Datum ondertekening machtiging ontbreekt.");
                errors.Add("Machtigingsdatum ontbreekt");
            }
        }

        string combinedDescription = TextProcessor.BuildDescription(generalDescription, record.DescriptionPart);
        if (string.IsNullOrWhiteSpace(combinedDescription))
        {
            messages.Add($"{rowPrefix} WAARSCHUWING: Omschrijving mag niet leeg zijn.");
            errors.Add("Omschrijving leeg");
        }
        else
        {
            if (combinedDescription.Length > 140)
            {
                messages.Add($"{rowPrefix} WAARSCHUWING: Omschrijving is langer dan 140 tekens.");
                errors.Add("Omschrijving te lang");
            }

            if (!TextProcessor.ContainsOnlyAllowedChars(combinedDescription))
            {
                messages.Add($"{rowPrefix} WAARSCHUWING: Omschrijving bevat ongeoorloofde tekens.");
                errors.Add("Omschrijving illegale tekens");
            }
        }

        if (!string.IsNullOrWhiteSpace(record.AddressLine1) && !TextProcessor.ContainsOnlyAllowedChars(record.AddressLine1))
        {
            messages.Add($"{rowPrefix} WAARSCHUWING: Adres1 bevat ongeoorloofde tekens.");
            errors.Add("Adres1 illegale tekens");
        }

        if (!string.IsNullOrWhiteSpace(record.AddressLine2) && !TextProcessor.ContainsOnlyAllowedChars(record.AddressLine2))
        {
            messages.Add($"{rowPrefix} WAARSCHUWING: Adres2 bevat ongeoorloofde tekens.");
            errors.Add("Adres2 illegale tekens");
        }

        if (!Regex.IsMatch(record.SequenceType, "^(FRST|RCUR|OOFF|FNAL)$", RegexOptions.CultureInvariant))
        {
            messages.Add($"{rowPrefix} WAARSCHUWING: Sequence type moet FRST/RCUR/OOFF/FNAL zijn.");
            errors.Add("Sequence type ongeldig");
        }

        return errors;
    }
}
