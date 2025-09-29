# SendGuard

SendGuard is an Outlook Add-in for Microsoft Office, built with C# targeting .NET Framework 4.8.

## Features

- Custom Outlook add-in functionality
- Policy-based email handling
- Integration with Outlook events

## Configuration

SendGuard uses a `policy.json` file to control allowed attachment types for specific recipient domains. This targets recipient email addresses at the entire domain, not individual email addresses.

**Config file locations:**
- Per-user: `%AppData%\SendGuard\policy.json`
- Machine-wide: `%ProgramData%\SendGuard\policy.json`

If both files exist, the per-user file takes precedence.

**Example `policy.json`:**

```
{
  "failSafeBlock": true,
  "targets": [
    { "domain": "recipient.com", "exts": [".gpg", ".pgp", ".asc"] },
    { "domain": "secure.example.au", "exts": [".gpg"] },
    { "domain": "*.partner.local", "exts": [".gpg", ".pgp"] }
  ]
}
```

This examples would set rules such that any email message going to any user at recipient.com, if there is are attachments, they must all have extensions of either `.gpg`, `.pgp`, or `.asc`. If there are attachments of any other extension, the message will be blocked. Messages without attachments are always allowed.

The extensions are not case sensitive.

If the configuration file is missing, malformed, or have empty rules, all messages are blocked.

After editing the configuration file, restart Outlook to apply changes.

## Getting Started

### Prerequisites

- Visual Studio 2022 or later
- .NET Framework 4.8
- Microsoft Office Outlook

### Building

1. Clone the repository.
2. Open `SendGuard.sln` in Visual Studio.
3. Restore NuGet packages.
4. Build the solution.

### Running

- Build and run the project in Visual Studio with Outlook installed.

## Project Structure

- `ThisAddIn.cs` - Main add-in logic
- `ThisAddIn.Designer.cs` - Designer-generated code
- `ThisAddIn.Designer.xml` - Designer metadata

## License

MIT License. See `LICENSE` file for details.
