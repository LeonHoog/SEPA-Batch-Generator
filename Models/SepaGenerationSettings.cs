namespace SEPA_Batch_Generator.Models
{
    public sealed class SepaGenerationSettings
    {
        public string CreditorName { get; init; } = string.Empty;
        public string CreditorIban { get; init; } = string.Empty;
        public string CreditorBic { get; init; } = string.Empty;
        public string CreditorId { get; init; } = string.Empty;
        public string GeneralDescription { get; init; } = string.Empty;
    }
}
