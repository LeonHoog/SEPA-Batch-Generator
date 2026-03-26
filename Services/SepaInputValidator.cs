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

            if (string.IsNullOrWhiteSpace(generalDescription) && importedRecords.All(r => string.IsNullOrWhiteSpace(r.DescriptionPart)))
            {
                messages.Add("Algemene omschrijving en regel-omschrijving mogen niet beide leeg zijn.");
            }

            foreach (var record in importedRecords)
            {
                var rowPrefix = $"Rij {record.RowNumber}:";
                var hasErrors = false;

                if (string.IsNullOrWhiteSpace(record.DebtorName))
                {
                    messages.Add($"{rowPrefix} Naam ontbreekt.");
                    hasErrors = true;
                }
                else if (!ContainsOnlyAllowedChars(record.DebtorName))
                {
                    messages.Add($"{rowPrefix} Naam bevat ongeoorloofde tekens.");
                    hasErrors = true;
                }

                if (!IsValidDutchIban(record.DebtorIban))
                {
                    messages.Add($"{rowPrefix} IBAN is ongeldig of niet Nederlands ({record.DebtorIban}).");
                    hasErrors = true;
                }

                if (record.Amount <= 0)
                {
                    messages.Add($"{rowPrefix} Bedrag moet groter zijn dan 0.");
                    hasErrors = true;
                }

                if (record.CollectionDate.Date < DateTime.Today)
                {
                    messages.Add($"{rowPrefix} Incassodatum ligt in het verleden ({record.CollectionDate:yyyy-MM-dd}).");
                    hasErrors = true;
                }

                if (string.IsNullOrWhiteSpace(record.MandateId))
                {
                    messages.Add($"{rowPrefix} Machtigingskenmerk ontbreekt.");
                    hasErrors = true;
                }

                if (record.MandateSignedOn == default)
                {
                    messages.Add($"{rowPrefix} Datum ondertekening machtiging ontbreekt.");
                    hasErrors = true;
                }

                var combinedDescription = BuildDescription(generalDescription, record.DescriptionPart);
                if (string.IsNullOrWhiteSpace(combinedDescription))
                {
                    messages.Add($"{rowPrefix} Omschrijving mag niet leeg zijn.");
                    hasErrors = true;
                }
                else
                {
                    if (combinedDescription.Length > 140)
                    {
                        messages.Add($"{rowPrefix} Omschrijving is langer dan 140 tekens.");
                        hasErrors = true;
                    }

                    if (!ContainsOnlyAllowedChars(combinedDescription))
                    {
                        messages.Add($"{rowPrefix} Omschrijving bevat ongeoorloofde tekens.");
                        hasErrors = true;
                    }
                }

                if (!string.IsNullOrWhiteSpace(record.AddressLine1) && !ContainsOnlyAllowedChars(record.AddressLine1))
                {
                    messages.Add($"{rowPrefix} Adres1 bevat ongeoorloofde tekens.");
                    hasErrors = true;
                }

                if (!string.IsNullOrWhiteSpace(record.AddressLine2) && !ContainsOnlyAllowedChars(record.AddressLine2))
                {
                    messages.Add($"{rowPrefix} Adres2 bevat ongeoorloofde tekens.");
                    hasErrors = true;
                }

                if (!Regex.IsMatch(record.SequenceType, "^(FRST|RCUR|OOFF|FNAL)$", RegexOptions.CultureInvariant))
                {
                    messages.Add($"{rowPrefix} Sequence type moet FRST/RCUR/OOFF/FNAL zijn.");
                    hasErrors = true;
                }

                if (!hasErrors)
                {
                    valid.Add(record);
                }
            }

            return valid;
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
