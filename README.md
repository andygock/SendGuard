# SendGuard

SendGuard is an Outlook Add-in for Microsoft Office, built with C# targeting .NET Framework 4.8. It is designed to prevent accidental sending of unencrypted email attachments of certain types, by enforcing policies on email attachments based on recipient domains.

## Disclaimer

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Features

- Custom Outlook add-in functionality
- Policy-based email handling
- Integration with Outlook events

## Configuration

SendGuard uses a `policy.json` file to control allowed attachment types for specific recipient domains. This targets recipient email addresses at the entire domain, not individual email addresses.

Config file locations:

- Per-user: `%AppData%\SendGuard\policy.json`
- Machine-wide: `%ProgramData%\SendGuard\policy.json`

If both files exist, the per-user file takes precedence.

Example `policy.json`:

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

## NSIS Installer

The project includes an NSIS-based installer script located at `NSIS/installer.nsi`. This script generates a Windows installer executable: `SendGuard-UserSetup.exe`.

### Installer Features

- Installs SendGuard add-in files to `%LocalAppData%\SendGuard`.
- Copies all necessary files from `bin/Release` (DLLs, .vsto, manifests).
- Ensures the target directory exists.
- Checks if Outlook is running and prompts the user to close it before installation.
- Displays a license agreement from `NSIS/license.txt`.
- Handles policy configuration files:
  - If `%AppData%\SendGuard\policy.json` does not exist, installs the default `policy.json`.
  - If a custom `policy.user.json` is present in the installer directory, it will be installed as `policy.json` for user-specific policies.
  - If `policy.json` already exists, it is not overwritten.

### Policy Files

- `NSIS/policy.json`: Default policy configuration installed if no user policy exists.
- `NSIS/policy.user.json`: Optional custom policy for user-specific installations.

### License

The installer displays the license from `NSIS/license.txt` during installation.

### Usage

Compile the plugin as Release in Visual Studio (not Debug).

To build the installer, use NSIS to compile `NSIS/installer.nsi`. The output will be `SendGuard-UserSetup.exe`.

To install:
1. Run `SendGuard-UserSetup.exe`.
2. Follow the prompts to accept the license and complete installation.
3. The installer will ensure Outlook is closed before proceeding.
4. Policy files will be handled as described above.

## Project Structure

- `ThisAddIn.cs` - Main add-in logic
- `ThisAddIn.Designer.cs` - Designer-generated code
- `ThisAddIn.Designer.xml` - Designer metadata
