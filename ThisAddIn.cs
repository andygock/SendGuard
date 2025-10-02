using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SendGuard
{
    public partial class ThisAddIn
    {
        private static Policy _policy = Policy.Default();
        private FileSystemWatcher _watcher;
        
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            LoadPolicy();
            this.Application.ItemSend += Application_ItemSend;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.ItemSend -= Application_ItemSend;
            if (_watcher != null) _watcher.Dispose();
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            var mail = item as Outlook.MailItem;
            if (mail == null) return;

            // Only block all sends if policy is in fail-safe mode OR has no rules
            if (_policy == null || _policy.Targets == null || _policy.Targets.Count == 0 ||
                (_policy.FailSafeBlock && _policy.Targets.All(t => t.Exts == null || t.Exts.Count == 0)))
            {
                MessageBox.Show(
                    "SendGuard is blocking all outgoing emails because the policy file is missing, malformed, or has no usable rules.\n\n" +
                    "Please edit your policy file to add allowed domains and extensions.",
                    "SendGuard Policy Enforcement", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cancel = true;
                return;
            }

            // Resolve SMTPs of all recipients
            var recips = new List<string>();
            foreach (Recipient r in mail.Recipients)
            {
                var smtp = SafeGetSmtp(r);
                if (!string.IsNullOrEmpty(smtp)) recips.Add(smtp.ToLowerInvariant());
            }
            if (recips.Count == 0) return;

            // Match any policy targets
            var hits = MatchingTargets(recips);
            if (hits.Count == 0) return;            // no target domain → allow
            if (mail.Attachments.Count == 0) return; // no attachments → allow

            // Allowed extensions are the union across matched targets
            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var t in hits) foreach (var e in t.Exts) allowed.Add(e);

            bool allOk = true;
            var bad = new List<string>();
            foreach (Outlook.Attachment a in mail.Attachments)
            {
                var name = (a.FileName ?? string.Empty);
                bool ok = allowed.Any(ext => name.EndsWith(ext, StringComparison.OrdinalIgnoreCase));
                if (!ok) { allOk = false; bad.Add(name); }
            }

            if (!allOk)
            {
                var hitList = string.Join(", ", hits.Select(h => h.Domain));
                var badList = string.Join("\n  • ", bad);
                MessageBox.Show(
                    "Send blocked.\n\nRecipients match policy: " + hitList +
                    "\nAllowed attachment extensions: " + string.Join(", ", allowed) +
                    "\n\nThese attachments are not allowed:\n  • " + badList +
                    "\n\nEncrypt to an approved format and try again.",
                    "SendGuard", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cancel = true;
            }
        }

        // ----- Policy load + watch -----

        private static readonly string UserPolicyPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SendGuard", "policy.json");
        //private static readonly string MachinePolicyPath =
        //    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "SendGuard", "policy.json");

        private static string PolicyPathInUse
        {
            get
            {
                // update: always use user path
                //if (File.Exists(UserPolicyPath)) return UserPolicyPath;
                //if (File.Exists(MachinePolicyPath)) return MachinePolicyPath;
                return UserPolicyPath; // Do not create the directory here
            }
        }

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
                        return; // Do not overwrite the file
                    }
                }

                if (p == null || p.Targets == null || p.Targets.Count == 0)
                {
                    p = new Policy
                    {
                        FailSafeBlock = true,
                        Targets = new List<Target>() // empty rules
                    };
                    if (!Directory.Exists(Path.GetDirectoryName(path)))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(path));
                    }
                    File.WriteAllText(path, p.ToJsonIndented());
                    MessageBox.Show(
                        "SendGuard policy file is missing, malformed, or has no rules.\n\n" +
                        $"A new policy file has been created at:\n{path}\n\n" +
                        "Please edit this file to add your allowed domains and extensions.\n\n" +
                        "Until this is done, all outgoing emails will be blocked.",
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
                    Targets = new List<Target>() // empty rules
                };
                File.WriteAllText(path, _policy.ToJsonIndented());
                MessageBox.Show(
                    $"SendGuard failed to load the policy file.\n\nError: {ex.Message}\n\n" +
                    $"A new policy file has been created at:\n{path}\n\n" +
                    "Please edit this file to add your allowed domains and extensions.\n\n" +
                    "Until this is done, all outgoing emails will be blocked.",
                    "SendGuard Policy Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void StartWatch()
        {
            var path = PolicyPathInUse;
            var dir = Path.GetDirectoryName(path);
            var file = Path.GetFileName(path);
            _watcher = new FileSystemWatcher(dir, file);
            _watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.CreationTime | NotifyFilters.FileName;
            _watcher.Changed += (s, e) => { try { LoadPolicy(); } catch { } };
            _watcher.Created += (s, e) => { try { LoadPolicy(); } catch { } };
            _watcher.Renamed += (s, e) => { try { LoadPolicy(); } catch { } };
            _watcher.EnableRaisingEvents = true;
        }

        private static List<Target> MatchingTargets(IEnumerable<string> recipientSmtps)
        {
            var hits = new List<Target>();
            foreach (var smtp in recipientSmtps)
            {
                var at = smtp.LastIndexOf('@');
                var dom = at >= 0 ? smtp.Substring(at + 1) : smtp;

                foreach (var t in _policy.Targets)
                {
                    if (t.Matches(dom)) hits.Add(t);
                }
            }
            // De-dup by domain string
            var uniq = new Dictionary<string, Target>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in hits) if (!uniq.ContainsKey(h.Domain)) uniq[h.Domain] = h;
            return uniq.Values.ToList();
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
            var logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SendGuard", "error.log");
            var logMessage = $"{DateTime.Now}: {message}\n{ex?.ToString()}\n\n";
            File.AppendAllText(logPath, logMessage);
        }
    }

    // ----- Policy models (Newtonsoft.Json) -----

    public class Policy
    {
        [JsonProperty("failSafeBlock")]
        public bool FailSafeBlock { get; set; }

        [JsonProperty("targets")]
        public List<Target> Targets { get; set; }

        public Policy()
        {
            FailSafeBlock = true;
            Targets = new List<Target>();
        }

        public void Normalise()
        {
            foreach (var t in Targets) t.Normalise();
            if (Targets.Count == 0 && FailSafeBlock)
            {
                Targets.Add(new Target { Domain = "*", Exts = new List<string>() });
            }
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
                Targets = new List<Target>
                {
                    new Target { Domain = "recipient.com", Exts = new List<string>{ ".gpg", ".pgp", ".asc" } }
                }
            };
        }
    }

    public class Target : IEquatable<Target>
    {
        [JsonProperty("domain")]
        public string Domain { get; set; } // "example.com", "*.example.com", or "*"

        [JsonProperty("exts")]
        public List<string> Exts { get; set; }

        [JsonIgnore]
        private Regex _wildcard;

        public Target()
        {
            Domain = string.Empty;
            Exts = new List<string>();
        }

        public void Normalise()
        {
            Domain = (Domain ?? string.Empty).Trim().ToLowerInvariant();
            var norm = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var e in (Exts ?? new List<string>())) norm.Add(e.Trim());
            Exts = norm.ToList();

            if (Domain == "*" || Domain.Length == 0)
            {
                _wildcard = new Regex("^.*$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            }
            else if (Domain.StartsWith("*.", StringComparison.Ordinal))
            {
                var root = Regex.Escape(Domain.Substring(2));
                _wildcard = new Regex(@"^[^.]+\." + root + "$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            }
            else
            {
                var exact = Regex.Escape(Domain);
                _wildcard = new Regex("^" + exact + "$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            }
        }

        public bool Matches(string domain)
        {
            return _wildcard != null && _wildcard.IsMatch(domain ?? string.Empty);
        }

        public bool Equals(Target other)
        {
            return other != null && string.Equals(Domain, other.Domain, StringComparison.OrdinalIgnoreCase);
        }

        public override int GetHashCode()
        {
            return Domain == null ? 0 : Domain.ToLowerInvariant().GetHashCode();
        }
    }
}
