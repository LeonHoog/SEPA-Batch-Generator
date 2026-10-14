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
            messages.Add(Localization.Get("ValidationWarningPrefix") + " " + Localization.Get("DescriptionsBothEmpty"));
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
            AddValidationError(messages, errors, rowPrefix, "ValidationNameMissing");
        else if (!TextProcessor.ContainsOnlyAllowedChars(record.DebtorName))
            AddValidationError(messages, errors, rowPrefix, "ValidationNameInvalidCharacters");

        if (!new IbanValidator().Validate(record.DebtorIban).IsValid)
            AddValidationError(messages, errors, rowPrefix, "ValidationIbanInvalid", record.DebtorIban);

        if (record.Amount <= 0)
            AddValidationError(messages, errors, rowPrefix, "ValidationAmountNonPositive", record.Amount);

        if (record.CollectionDate.Date < DateTime.Today)
            AddValidationError(messages, errors, rowPrefix, "ValidationCollectionDatePast", record.CollectionDate.ToString("yyyy-MM-dd"));

        bool mandateIdMissing = string.IsNullOrWhiteSpace(record.MandateId);
        bool mandateDateMissing = record.MandateSignedOn == default;

        if (mandateIdMissing || mandateDateMissing)
        {
            if (mandateIdMissing && mandateDateMissing)
                AddValidationError(messages, errors, rowPrefix, "ValidationMandateAndDateMissing");
            else if (mandateIdMissing)
                AddValidationError(messages, errors, rowPrefix, "ValidationMandateMissing");
            else
                AddValidationError(messages, errors, rowPrefix, "ValidationMandateDateMissing");
        }

        string combinedDescription = TextProcessor.BuildDescription(generalDescription, record.DescriptionPart);
        if (string.IsNullOrWhiteSpace(combinedDescription))
        {
            AddValidationError(messages, errors, rowPrefix, "ValidationDescriptionEmpty");
        }
        else
        {
            if (combinedDescription.Length > 140)
                AddValidationError(messages, errors, rowPrefix, "ValidationDescriptionTooLong");

            if (!TextProcessor.ContainsOnlyAllowedChars(combinedDescription))
                AddValidationError(messages, errors, rowPrefix, "ValidationDescriptionInvalidCharacters");
        }

        if (!string.IsNullOrWhiteSpace(record.AddressLine1) && !TextProcessor.ContainsOnlyAllowedChars(record.AddressLine1))
            AddValidationError(messages, errors, rowPrefix, "ValidationAddress1InvalidCharacters");

        if (!string.IsNullOrWhiteSpace(record.AddressLine2) && !TextProcessor.ContainsOnlyAllowedChars(record.AddressLine2))
            AddValidationError(messages, errors, rowPrefix, "ValidationAddress2InvalidCharacters");

        if (!Regex.IsMatch(record.SequenceType, "^(FRST|RCUR|OOFF|FNAL)$", RegexOptions.CultureInvariant))
            AddValidationError(messages, errors, rowPrefix, "ValidationSequenceTypeInvalid");

        return errors;
    }

    private static void AddValidationError(List<string> messages, List<string> errors, string rowPrefix, string resourceKey, params object[] args)
    {
        string reason = Localization.Get(resourceKey, args);
        messages.Add($"{rowPrefix} {Localization.Get("ValidationWarningPrefix")} {reason}");
        errors.Add(reason);
    }
}
