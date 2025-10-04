# SendGuard

SendGuard is an Outlook Add-in for Microsoft Office, built with C# targeting .NET Framework 4.8. It is designed to prevent accidental sending of sensitive email attachments by enforcing a configurable set of rules that match recipient email addresses and attachment filenames.

## Disclaimer

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Features

- Policy-based email attachment filtering
- Line-by-line rule processing with "first match wins" logic
- Glob wildcard support (`*`, `?`) for flexible matching of recipients and attachments
- Case-insensitive matching for all rules
- Fail-safe mode to block any attachments not explicitly allowed
- Seamless integration with the Outlook `ItemSend` event

## Configuration

SendGuard's behavior is controlled by a `policy.json` file. This file contains a list of rules that are processed sequentially from top to bottom. For every recipient and every attachment on an outgoing email, SendGuard checks the rules in order. **The first rule that matches the recipient/attachment pair determines the outcome.**

- If the first matching rule has `"accept": true`, that specific attachment is allowed for that recipient.
- If the first matching rule has `"accept": false`, the attachment is blocked, and the entire email is prevented from sending.
- If an attachment/recipient pair does not match any rule, its fate is determined by the `failSafeBlock` setting.

Config file location:

- Per-user: `%AppData%\SendGuard\policy.json`

### New `policy.json` Format

The policy file now consists of a `failSafeBlock` flag and a list of `rules`.

- **`failSafeBlock`**: (boolean) If `true` (recommended default), any attachment/recipient combination that does not match an explicit `accept: true` rule will be blocked. This prevents accidental sends if your rules are not comprehensive. If `false`, anything not matching a rule is allowed.
- **`rules`**: (array) An ordered list of rule objects. Each rule has three properties:
  - **`to`**: A glob pattern to match against the recipient's full SMTP email address (e.g., `user@example.com`, `*@example.com`, `*@*.partner.local`).
  - **`attachment`**: A glob pattern to match against the attachment's full filename (e.g., `report.docx`, `*.gpg`, `confidential-*.zip`).
  - **`accept`**: A boolean value (`true` to allow, `false` to block).

### Example `policy.json`

```json
{
  "failSafeBlock": true,
  "rules": [
    // Rule 1 & 2: Allow specific encrypted file types to an external partner domain.
    { "to": "*@partner.com", "attachment": "*.gpg", "accept": true },
    { "to": "*@partner.com", "attachment": "*.pgp", "accept": true },

    // Rule 3: Block ALL other attachment types to that partner. This is a crucial catch-all.
    { "to": "*@partner.com", "attachment": "*", "accept": false },

    // Rule 4: Allow sending any ZIP file to the internal 'archive' team.
    { "to": "archive@mycompany.local", "attachment": "*.zip", "accept": true },

    // Rule 5: Explicitly block sending executables to anyone, anywhere.
    // This rule is placed high to ensure it's checked before more permissive rules.
    { "to": "*", "attachment": "*.exe", "accept": false },
    { "to": "*", "attachment": "*.dll", "accept": false },

    // Rule 6: Allow sending any document to anyone within the company.
    { "to": "*@mycompany.local", "attachment": "*.docx", "accept": true },
    { "to": "*@mycompany.local", "attachment": "*.xlsx", "accept": true },
    { "to": "*@mycompany.local", "attachment": "*.pdf", "accept": true }
  ]
}
```

### How It Works: An Example

Imagine you send an email with two attachments, `report.docx` and `installer.exe`, to `jane@mycompany.local` and `bob@partner.com`.

1.  **For `bob@partner.com` and `installer.exe`**:
    - The add-in checks the rules.
    - Rule 5 (`"to": "*", "attachment": "*.exe", "accept": false`) is the first match.
    - The action is `accept: false`.
    - **The entire email is immediately blocked.** A message box appears explaining the violation. The send operation is cancelled.

2.  If you remove `installer.exe` and try to send only `report.docx` to both recipients:
    - **For `bob@partner.com` and `report.docx`**:
        - Rules 1 and 2 are skipped (filename doesn't match `*.gpg` or `*.pgp`).
        - Rule 3 (`"to": "*@partner.com", "attachment": "*", "accept": false`) is a match.
        - **The email is blocked.**
    - **For `jane@mycompany.local` and `report.docx`**:
        - Rules 1-5 are skipped.
        - Rule 6 (`"to": "*@mycompany.local", "attachment": "*.docx", "accept": true`) is a match.
        - This combination is allowed.

However, because the combination for `bob@partner.com` was blocked, the entire send operation fails. **All attachments must be allowed for all recipients.**

### Important Notes

- Emails without attachments are **always allowed** and are not processed by SendGuard.
- All matching (`to` and `attachment`) is **case-insensitive**.
- If the configuration file is missing, malformed, or has empty rules, all messages with attachments will be blocked (assuming default `failSafeBlock: true`). A new, empty policy file will be created.
- After editing `policy.json`, you must **restart Outlook** for the changes to take effect.

## Getting Started

### Prerequisites

- Visual Studio 2022 or later
- .NET Framework 4.8
- Microsoft Office Outlook

### Building

- Clone the repository.
- Open `SendGuard.sln` in Visual Studio.
- Restore NuGet packages if necessary.
- Build the solution.

### Running

- Build and run the project from Visual Studio with Outlook installed. This will launch Outlook with the add-in attached for debugging.

## NSIS Installer

The project includes an NSIS-based installer script at `NSIS/installer.nsi`. This script generates a Windows installer executable: `SendGuard-UserSetup.exe`.

### Installer Features

- Installs SendGuard add-in files to `%LocalAppData%\SendGuard`.
- Copies all necessary files from `bin/Release` (DLLs, .vsto, manifests).
- Checks if Outlook is running and prompts the user to close it before installation.
- Displays a license agreement from `NSIS/license.txt`.
- Handles policy configuration files:
  - If `%AppData%\SendGuard\policy.json` does not exist, it installs the default `NSIS/policy.json`.
  - If a custom `policy.user.json` is present in the installer directory, it will be installed as `policy.json`.
  - **If `policy.json` already exists, it is NOT overwritten.**

### Policy Files

- `NSIS/policy.json`: A default policy configuration. It's recommended to ship this with an empty `rules` array and `failSafeBlock: true` to ensure a secure-by-default state, forcing the user to create their own rules.
- `NSIS/policy.user.json`: An optional, pre-configured policy file that can be bundled with the installer for specific deployments.

### Building the Installer

- Build the project in **Release** configuration within Visual Studio.
- Install [NSIS (Nullsoft Scriptable Install System)](https://nsis.sourceforge.io/Download).
- Right-click `NSIS/installer.nsi` and select "Compile NSIS Script".
- The output, `SendGuard-UserSetup.exe`, will be created.

## Project Structure

- `ThisAddIn.cs` - Main add-in logic, event handling, and policy enforcement.
- `ThisAddIn.Designer.cs` - Designer-generated code.
- `ThisAddIn.Designer.xml` - Designer metadata.