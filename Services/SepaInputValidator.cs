using SEPA_Batch_Generator.Models;
using System.Text;
using System.Text.RegularExpressions;

namespace SEPA_Batch_Generator.Services
{
    public sealed class SepaInputValidator
    {
        private static readonly HashSet<char> AllowedCharacters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/-?:().,'+ ".ToHashSet();

        public static List<DirectDebitRecord> Validate(List<DirectDebitRecord> importedRecords, string generalDescription, List<string> messages)
        {
            var valid = new List<DirectDebitRecord>();
            var rejected = new List<(DirectDebitRecord Record, string Reason)>();

            if (string.IsNullOrWhiteSpace(generalDescription) && importedRecords.All(r => string.IsNullOrWhiteSpace(r.DescriptionPart)))
            {
                messages.Add("WAARSCHUWING: Algemene omschrijving en regel-omschrijving mogen niet beide leeg zijn.");
            }

            foreach (var record in importedRecords)
            {
                var rowPrefix = $"Rij {record.RowNumber}:";
                var hasErrors = false;
                var errorReasons = new List<string>();

                if (string.IsNullOrWhiteSpace(record.DebtorName))
                {
                    messages.Add($"{rowPrefix} WAARSCHUWING: Naam ontbreekt.");
                    hasErrors = true;
                    errorReasons.Add("Naam ontbreekt");
                }
                else if (!ContainsOnlyAllowedChars(record.DebtorName))
                {
                    messages.Add($"{rowPrefix} WAARSCHUWING: Naam bevat ongeoorloofde tekens.");
                    hasErrors = true;
                    errorReasons.Add("Naam bevat ongeoorloofde tekens");
                }

                if (!IsValidDutchIban(record.DebtorIban))
                {
                    messages.Add($"{rowPrefix} WAARSCHUWING: IBAN is ongeldig of niet Nederlands ({record.DebtorIban}).");
                    hasErrors = true;
                    errorReasons.Add("IBAN ongeldig");
                }

                if (record.Amount <= 0)
                {
                    messages.Add($"{rowPrefix} WAARSCHUWING: Bedrag mag niet 0 of negatief zijn (bedrag: {record.Amount}).");
                    hasErrors = true;
                    errorReasons.Add("Bedrag 0 of negatief");
                }

                if (record.CollectionDate.Date < DateTime.Today)
                {
                    messages.Add($"{rowPrefix} WAARSCHUWING: Incassodatum ligt in het verleden ({record.CollectionDate:yyyy-MM-dd}).");
                    hasErrors = true;
                    errorReasons.Add("Incassodatum in verleden");
                }

                var mandateIdMissing = string.IsNullOrWhiteSpace(record.MandateId);
                var mandateDateMissing = record.MandateSignedOn == default;

                if (mandateIdMissing || mandateDateMissing)
                {
                    if (mandateIdMissing && mandateDateMissing)
                    {
                        messages.Add($"{rowPrefix} WAARSCHUWING: Zowel machtigingskenmerk als ondertekeningsdatum ontbreken.");
                        errorReasons.Add("Machtigingskenmerk en -datum ontbreken");
                    }
                    else if (mandateIdMissing)
                    {
                        messages.Add($"{rowPrefix} WAARSCHUWING: Machtigingskenmerk ontbreekt.");
                        errorReasons.Add("Machtigingskenmerk ontbreekt");
                    }
                    else
                    {
                        messages.Add($"{rowPrefix} WAARSCHUWING: Datum ondertekening machtiging ontbreekt.");
                        errorReasons.Add("Machtigingsdatum ontbreekt");
                    }
                    hasErrors = true;
                }

                var combinedDescription = BuildDescription(generalDescription, record.DescriptionPart);
                if (string.IsNullOrWhiteSpace(combinedDescription))
                {
                    messages.Add($"{rowPrefix} WAARSCHUWING: Omschrijving mag niet leeg zijn.");
                    hasErrors = true;
                    errorReasons.Add("Omschrijving leeg");
                }
                else
                {
                    if (combinedDescription.Length > 140)
                    {
                        messages.Add($"{rowPrefix} WAARSCHUWING: Omschrijving is langer dan 140 tekens.");
                        hasErrors = true;
                        errorReasons.Add("Omschrijving te lang");
                    }

                    if (!ContainsOnlyAllowedChars(combinedDescription))
                    {
                        messages.Add($"{rowPrefix} WAARSCHUWING: Omschrijving bevat ongeoorloofde tekens.");
                        hasErrors = true;
                        errorReasons.Add("Omschrijving illegale tekens");
                    }
                }

                if (!string.IsNullOrWhiteSpace(record.AddressLine1) && !ContainsOnlyAllowedChars(record.AddressLine1))
                {
                    messages.Add($"{rowPrefix} WAARSCHUWING: Adres1 bevat ongeoorloofde tekens.");
                    hasErrors = true;
                    errorReasons.Add("Adres1 illegale tekens");
                }

                if (!string.IsNullOrWhiteSpace(record.AddressLine2) && !ContainsOnlyAllowedChars(record.AddressLine2))
                {
                    messages.Add($"{rowPrefix} WAARSCHUWING: Adres2 bevat ongeoorloofde tekens.");
                    hasErrors = true;
                    errorReasons.Add("Adres2 illegale tekens");
                }

                if (!Regex.IsMatch(record.SequenceType, "^(FRST|RCUR|OOFF|FNAL)$", RegexOptions.CultureInvariant))
                {
                    messages.Add($"{rowPrefix} WAARSCHUWING: Sequence type moet FRST/RCUR/OOFF/FNAL zijn.");
                    hasErrors = true;
                    errorReasons.Add("Sequence type ongeldig");
                }

                if (hasErrors)
                {
                    rejected.Add((record, string.Join("; ", errorReasons)));
                }
                else
                {
                    valid.Add(record);
                }
            }

            // Store rejected records for reporting
            _lastRejectedRecords = rejected;

            return valid;
        }

        private static List<(DirectDebitRecord Record, string Reason)> _lastRejectedRecords = [];

        public static List<(DirectDebitRecord Record, string Reason)> GetLastRejectedRecords()
        {
            return _lastRejectedRecords;
        }

        public static string BuildDescription(string generalDescription, string? rowDescription)
        {
            return string.Join(' ', new[] { generalDescription?.Trim(), rowDescription?.Trim() }.Where(v => !string.IsNullOrWhiteSpace(v)));
        }

        private static bool ContainsOnlyAllowedChars(string value)
        {
            return value.All(c => AllowedCharacters.Contains(c));
        }

        private static bool IsValidDutchIban(string iban)
        {
            if (string.IsNullOrWhiteSpace(iban))
            {
                return false;
            }

            var normalized = iban.Replace(" ", string.Empty).ToUpperInvariant();
            if (!Regex.IsMatch(normalized, "^NL[0-9]{2}[A-Z]{4}[0-9]{10}$", RegexOptions.CultureInvariant))
            {
                return false;
            }

            var rearranged = normalized[4..] + normalized[..4];
            var numeric = new StringBuilder();
            foreach (var c in rearranged)
            {
                if (char.IsLetter(c))
                {
                    numeric.Append(c - 'A' + 10);
                }
                else
                {
                    numeric.Append(c);
                }
            }

            var remainder = 0;
            foreach (var c in numeric.ToString())
            {
                remainder = (remainder * 10 + (c - '0')) % 97;
            }

            return remainder == 1;
        }
    }
}
