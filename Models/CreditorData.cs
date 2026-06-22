namespace SEPA_Batch_Generator.Models;

public sealed class CreditorData : IEquatable<CreditorData>
{
    public string Name { get; private set; } = string.Empty;
    public string Iban { get; private set; } = string.Empty;
    public string Bic { get; private set; } = string.Empty;
    public string Id { get; private set; } = string.Empty;

    public CreditorData() { }

    public CreditorData(string name, string iban, string bic, string id)
    {
        Name = name?.Trim() ?? string.Empty;
        Iban = (iban?.Trim() ?? string.Empty).Replace(" ", string.Empty).ToUpperInvariant();
        Bic = (bic?.Trim() ?? string.Empty).ToUpperInvariant();
        Id = id?.Trim() ?? string.Empty;
    }

    public bool IsComplete => !string.IsNullOrWhiteSpace(Name) && 
                               !string.IsNullOrWhiteSpace(Iban) && 
                               !string.IsNullOrWhiteSpace(Bic) && 
                               !string.IsNullOrWhiteSpace(Id);

    public bool Equals(CreditorData? other)
    {
        if (other is null) return false;
        return (Name == other.Name) && (Iban == other.Iban) && (Bic == other.Bic) && (Id == other.Id);
    }

    public override bool Equals(object? obj) => obj is CreditorData other && Equals(other);
    public override int GetHashCode() => HashCode.Combine(Name, Iban, Bic, Id);
    public override string ToString() => $"{Name} ({Iban})";
}
