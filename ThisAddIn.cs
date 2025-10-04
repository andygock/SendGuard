using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SendGuard
{
    public partial class ThisAddIn
    {
        private static Policy _policy = Policy.Default();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            LoadPolicy();
            this.Application.ItemSend += Application_ItemSend;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.ItemSend -= Application_ItemSend;
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            var mail = item as Outlook.MailItem;
            if (mail == null) return;

            // If there are no attachments, the policy rules do not apply, so allow the send.
            if (mail.Attachments.Count == 0) return;

            // Fail-safe check: if policy is missing or has no rules in fail-safe mode, block all emails with attachments.
            if (_policy == null || (_policy.FailSafeBlock && (_policy.Rules == null || _policy.Rules.Count == 0)))
            {
                MessageBox.Show(
                    "SendGuard is blocking this email because the policy file is missing or has no rules in fail-safe mode.\n\n" +
                    "Please edit your policy file to add rules for sending attachments.",
                    "SendGuard Policy Enforcement", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cancel = true;
                return;
            }

            // Resolve SMTPs of all recipients
            var recipients = new List<string>();
            foreach (Recipient r in mail.Recipients)
            {
                var smtp = SafeGetSmtp(r);
                if (!string.IsNullOrEmpty(smtp)) recipients.Add(smtp.ToLowerInvariant());
            }
            if (recipients.Count == 0) return;

            // Check each attachment against each recipient based on the ordered rules.
            foreach (Outlook.Attachment attachment in mail.Attachments)
            {
                var attachmentName = attachment.FileName ?? string.Empty;
                foreach (var recipient in recipients)
                {
                    bool? isAllowed = null;
                    bool ruleMatched = false;

                    // Find the first matching rule and get its decision.
                    foreach (var rule in _policy.Rules)
                    {
                        if (rule.Matches(recipient, attachmentName))
                        {
                            isAllowed = rule.Accept;
                            ruleMatched = true;
                            break; // First match wins.
                        }
                    }

                    bool block = false;
                    string reason = "";

                    if (ruleMatched)
                    {
                        // An explicit rule was found. Block if it's an "accept: false" rule.
                        if (!isAllowed.Value) // isAllowed will have a value if ruleMatched is true.
                        {
                            block = true;
                            reason = $"Attachment '{attachmentName}' is explicitly blocked for recipient '{recipient}'.";
                        }
                    }
                    else // No rule matched this combination.
                    {
                        // If in fail-safe mode, block anything not explicitly allowed.
                        if (_policy.FailSafeBlock)
                        {
                            block = true;
                            reason = $"Attachment '{attachmentName}' is not covered by any rule for recipient '{recipient}' (fail-safe mode).";
                        }
                    }

                    if (block)
                    {
                        MessageBox.Show(
                            "Send blocked due to a policy violation.\n\n" +
                            reason +
                            "\n\nReview your attachments or recipients and try again.",
                            "SendGuard", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        cancel = true;
                        return; // One violation is enough to stop the entire send.
                    }
                }
            }
            // If we get here, all attachment/recipient pairs are allowed.
        }

        // ----- Policy load + watch -----

        private static readonly string UserPolicyPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SendGuard", "policy.json");

        private static string PolicyPathInUse => UserPolicyPath;

        private void LoadPolicy()
        {
            try
            {
                var path = PolicyPathInUse;
                Policy p = null;
                if (File.Exists(path))
                {
                    var json = File.ReadAllText(path);
                    try
                    {
                        p = JsonConvert.DeserializeObject<Policy>(json);
                    }
                    catch (JsonException ex)
                    {
                        MessageBox.Show(
                            $"Failed to parse policy file at:\n{path}\n\nError: {ex.Message}\n\n" +
                            "The policy file will not be overwritten. Please fix the file manually.",
                            "SendGuard Policy Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // Do not overwrite the corrupted file.
                    }
                }

                if (p == null || p.Rules == null)
                {
                    p = new Policy
                    {
                        FailSafeBlock = true,
                        Rules = new List<Rule>() // empty rules
                    };
                    var dir = Path.GetDirectoryName(path);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }
                    File.WriteAllText(path, p.ToJsonIndented());
                    MessageBox.Show(
                        "SendGuard policy file is missing, malformed, or has no rules.\n\n" +
                        $"A new, empty policy file has been created at:\n{path}\n\n" +
                        "Please edit this file to add your rules for sending attachments.\n\n" +
                        "Because 'failSafeBlock' is true, all outgoing emails with attachments will be blocked until rules are added.",
                        "SendGuard Policy Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                p.Normalise();
                _policy = p;
            }
            catch (System.Exception ex)
            {
                var path = PolicyPathInUse;
                _policy = new Policy
                {
                    FailSafeBlock = true,
                    Rules = new List<Rule>() // empty rules
                };
                if (!string.IsNullOrEmpty(path))
                {
                    File.WriteAllText(path, _policy.ToJsonIndented());
                }
                MessageBox.Show(
                    $"SendGuard failed to load or create the policy file.\n\nError: {ex.Message}\n\n" +
                    "Until the issue is resolved, all outgoing emails with attachments will be blocked.",
                    "SendGuard Policy Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogError("Failed to load or create policy file.", ex);
            }
        }

        // Resolve reliable SMTP per Microsoft interop guidance
        private static string SafeGetSmtp(Recipient r)
        {
            try
            {
                var ae = r?.AddressEntry;
                if (ae == null) return null;

                var exUser = ae.GetExchangeUser();
                if (exUser != null && !string.IsNullOrEmpty(exUser.PrimarySmtpAddress))
                    return exUser.PrimarySmtpAddress;

                var exList = ae.GetExchangeDistributionList();
                if (exList != null && !string.IsNullOrEmpty(exList.PrimarySmtpAddress))
                    return exList.PrimarySmtpAddress;

                if (string.Equals(ae.Type, "SMTP", StringComparison.OrdinalIgnoreCase))
                    return ae.Address;

                const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                var pa = ae.PropertyAccessor;
                return pa?.GetProperty(PR_SMTP_ADDRESS) as string;
            }
            catch (System.Exception ex)
            {
                LogError("Failed to resolve SMTP address for recipient.", ex);
                return null;
            }
        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        private static void LogError(string message, System.Exception ex = null)
        {
            try
            {
                var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SendGuard");
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                var logPath = Path.Combine(logDir, "error.log");
                var logMessage = $"{DateTime.Now}: {message}\n{ex?.ToString()}\n\n";
                File.AppendAllText(logPath, logMessage);
            }
            catch
            {
                // Ignore errors during logging
            }
        }
    }

    // ----- Policy models (Newtonsoft.Json) -----

    public class Policy
    {
        [JsonProperty("failSafeBlock")]
        public bool FailSafeBlock { get; set; }

        [JsonProperty("rules")]
        public List<Rule> Rules { get; set; }

        public Policy()
        {
            FailSafeBlock = true;
            Rules = new List<Rule>();
        }

        public void Normalise()
        {
            if (Rules == null) Rules = new List<Rule>();
            foreach (var r in Rules) r.Normalise();
        }

        public string ToJsonIndented()
        {
            return JsonConvert.SerializeObject(this, Formatting.Indented,
                new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
        }

        public static Policy Default()
        {
            return new Policy
            {
                FailSafeBlock = true,
                Rules = new List<Rule>()
            };
        }
    }

    public class Rule
    {
        [JsonProperty("to")]
        public string To { get; set; }

        [JsonProperty("attachment")]
        public string Attachment { get; set; }

        [JsonProperty("accept")]
        public bool Accept { get; set; }

        [JsonIgnore]
        private Regex _toRegex;
        [JsonIgnore]
        private Regex _attachmentRegex;

        public Rule()
        {
            To = string.Empty;
            Attachment = string.Empty;
        }

        public void Normalise()
        {
            To = (To ?? string.Empty).Trim();
            Attachment = (Attachment ?? string.Empty).Trim();

            _toRegex = new Regex(GlobToRegex(To), RegexOptions.Compiled | RegexOptions.IgnoreCase);
            _attachmentRegex = new Regex(GlobToRegex(Attachment), RegexOptions.Compiled | RegexOptions.IgnoreCase);
        }

        public bool Matches(string email, string attachmentName)
        {
            if (_toRegex == null || _attachmentRegex == null) return false;

            return _toRegex.IsMatch(email ?? string.Empty) &&
                   _attachmentRegex.IsMatch(attachmentName ?? string.Empty);
        }

        private static string GlobToRegex(string glob)
        {
            if (string.IsNullOrEmpty(glob))
            {
                return "^$"; // Match empty string exactly
            }

            // 1. Escape all special regex characters.
            var regex = new StringBuilder(Regex.Escape(glob));

            // 2. Un-escape the glob wildcards back to their regex equivalents.
            regex.Replace(@"\*", ".*");
            regex.Replace(@"\?", ".");

            // 3. Anchor the pattern to match the whole string.
            return $"^{regex}$";
        }
    }
}