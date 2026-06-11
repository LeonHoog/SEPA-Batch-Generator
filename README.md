<table style="border: none; border-collapse: collapse;">
    <tr style="border: none;">
        <td style="border: none; padding-right: 20px;">
            <h1>SEPA Batch Generator</h1>
            Generate SEPA XML (pain.008) batch files for direct debits using spreadsheet input.
        </td>
        <td style="border: none;">
            <img width="200" src="https://raw.githubusercontent.com/LeonHoog/SEPA-Batch-Generator/refs/heads/main/Assets/logo.ico" alt="SEPA Batch Generator Logo"></td>
    </tr>
</table>


## Installation
Clone the repository and build with the .NET SDK, or check for the latest distributed binary under [Releases](https://github.com/LeonHoog/SEPA-Batch-Generator/releases).

Build from source (example):
```bash
git clone https://github.com/LeonHoog/SEPA-Batch-Generator.git
```

Run the app _(from within the project folder)_:
```bash
dotnet run -c Release
```

## Usage
1. Open the application.
2. Load an Excel file containing your direct-debit records.
3. Configure creditor settings (BIC, IBAN, creditor ID, defaults) in the settings UI, which are stored in `creditor.ini`.
4. Map spreadsheet columns to the expected fields using the column-mapping UI.
5. Run validation (and check for flagged errors/warnings).
6. Generate the SEPA XML batch and save it for submission to your bank.

## Input format
- Each row represents a single direct-debit instruction.
- Required fields include: debtor name, IBAN, amount, mandate ID & mandate date.
- Use the app's column mapping section to align columns from your spreadsheet to the required fields.

## Configuration
- `creditor.ini` contains creditor-specific defaults.
- Application settings are persisted in `settings.ini`.

## Contributing
Contributions are welcome. Please open issues or pull requests on the repository.
