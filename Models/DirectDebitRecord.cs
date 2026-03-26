namespace SEPA_Batch_Generator.Models
{
    public sealed class DirectDebitRecord
    {
        public int RowNumber { get; init; }
        public string DebtorName { get; init; } = string.Empty;
        public string DebtorIban { get; init; } = string.Empty;
        public string? DebtorBic { get; init; }
        public decimal Amount { get; init; }
        public string MandateId { get; init; } = string.Empty;
        public DateTime MandateSignedOn { get; init; }
        public DateTime CollectionDate { get; init; }
        public string SequenceType { get; init; } = "RCUR";
        public string? DescriptionPart { get; init; }
        public string? AddressLine1 { get; init; }
        public string? AddressLine2 { get; init; }
    }
}
