namespace SEPA_Batch_Generator.Models
{
    public sealed class ExcelLayoutSettings
    {
        public string DebtorNameColumn { get; init; } = string.Empty;
        public string? DebtorLastNameColumn { get; init; }
        public string DebtorIbanColumn { get; init; } = string.Empty;
        public string? DebtorBicColumn { get; init; }
        public string AmountColumn { get; init; } = string.Empty;
        public string MandateIdColumn { get; init; } = string.Empty;
        public string MandateDateColumn { get; init; } = string.Empty;
        public string? CollectionDateColumn { get; init; }
        public string SequenceTypeColumn { get; init; } = string.Empty;
        public string DescriptionColumn { get; init; } = string.Empty;
        public string? Address1Column { get; init; }
        public string? Address2Column { get; init; }
    }
}
