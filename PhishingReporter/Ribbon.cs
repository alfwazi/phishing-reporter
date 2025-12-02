/*
 * Developer: Abdulla Albreiki
 * Github: https://github.com/0dteam
 * licensed under the GNU General Public License v3.0
 */
 
using Microsoft.Office.Core;
using PhishingReporter.Properties;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Security.Cryptography;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Text;



namespace PhishingReporter
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public Bitmap getGroup1Image(IRibbonControl control)
        {
            return Resources.phishing;
        }

        /// <summary>
        /// Main function to report phishing emails
        /// </summary>
        public void reportPhishing(Office.IRibbonControl control)
        {
            try
            {
                var result = MessageBox.Show(
                    "Do you want to report this email to geidea Cybersecurity team as a potential phishing attempt?",
                    "Report Phishing Email",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);

                if (result == DialogResult.Yes)
                {
                    reportPhishingEmailToSecurityTeam(control);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    $"An error occurred: {ex.Message}\n\nPlease contact support if this persists.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /*
         *  Helper functions 
         */

        /// <summary>
        /// Reports phishing email to security team with enhanced error handling
        /// </summary>
        private void reportPhishingEmailToSecurityTeam(IRibbonControl control)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null)
                {
                    MessageBox.Show("Unable to access Outlook. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Selection selection = explorer.Selection;
                if (selection == null || selection.Count < 1)
                {
                    MessageBox.Show("Please select an email before reporting.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (selection.Count > 1)
                {
                    MessageBox.Show("You can only report one email at a time. Please select a single email.", "Multiple Items Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Get the selected item
                object selectedItem = selection[1];
                if (selectedItem == null)
                {
                    MessageBox.Show("Unable to access the selected item. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Identify item type
                string reportedItemType = GetItemType(selectedItem);
                if (string.IsNullOrEmpty(reportedItemType))
                {
                    MessageBox.Show("This item type cannot be reported. Please select an email, meeting, contact, appointment, or task.", "Invalid Item Type", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                MailItem mailItem = (selectedItem as MailItem);
                string reportedItemHeaders = GetEmailHeaders(mailItem, reportedItemType);

                // Check if this is a simulated phishing campaign
                string simulatedPhishingURL = GoPhishIntegration.SetReportURL(reportedItemHeaders);

                if (!string.IsNullOrEmpty(simulatedPhishingURL))
                {
                    // This is a GoPhish simulated campaign
                    ProcessGoPhishCampaign(simulatedPhishingURL, mailItem);
                }
                else
                {
                    // This is a real suspicious email
                    ProcessSuspiciousEmail(selectedItem, mailItem, reportedItemType, reportedItemHeaders);
                }
            }
            catch (COMException comEx)
            {
                HandleError("Outlook communication error", comEx);
            }
            catch (System.Exception ex)
            {
                HandleError("Unexpected error occurred", ex);
            }
        }

        /// <summary>
        /// Gets the type of the selected Outlook item
        /// </summary>
        private string GetItemType(object item)
        {
            if (item is MailItem) return "MailItem";
            if (item is MeetingItem) return "MeetingItem";
            if (item is ContactItem) return "ContactItem";
            if (item is AppointmentItem) return "AppointmentItem";
            if (item is TaskItem) return "TaskItem";
            return null;
        }

        /// <summary>
        /// Extracts email headers if available
        /// </summary>
        private string GetEmailHeaders(MailItem mailItem, string itemType)
        {
            if (itemType == "MailItem" && mailItem != null)
            {
                try
                {
                    return mailItem.HeaderString() ?? "Headers not available";
                }
                catch
                {
                    return "Headers could not be extracted";
                }
            }
            return $"Headers not available - Item type: {itemType}";
        }

        /// <summary>
        /// Processes a GoPhish simulated phishing campaign
        /// </summary>
        private void ProcessGoPhishCampaign(string reportURL, MailItem mailItem)
        {
            try
            {
                bool success = GoPhishIntegration.SendReportNotificationToServer(reportURL);
                
                if (success)
                {
                    Properties.Settings.Default.gophish_reports_counter++;
                    Properties.Settings.Default.Save();

                    MessageBox.Show(
                        "✅ Excellent! You've successfully reported a simulated phishing campaign.\n\n" +
                        "Your security awareness is helping protect geidea!",
                        "Phishing Simulation Reported",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    // Delete the email after successful reporting
                    if (mailItem != null)
                    {
                        mailItem.Delete();
                    }
                }
                else
                {
                    MessageBox.Show(
                        "⚠️ The GoPhish server could not be reached, but your report has been recorded.\n\n" +
                        "The email will not be deleted. Please try again or contact support.",
                        "Connection Issue",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (System.Exception ex)
            {
                HandleError("Error processing GoPhish campaign", ex);
            }
        }

        /// <summary>
        /// Processes a real suspicious email report
        /// </summary>
        private void ProcessSuspiciousEmail(object selectedItem, MailItem mailItem, string itemType, string headers)
        {
            try
            {
                MailItem reportEmail = Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem) as MailItem;
                if (reportEmail == null)
                {
                    MessageBox.Show("Unable to create report email. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Set recipient
                reportEmail.To = Properties.Settings.Default.infosec_email ?? "SOC <SOC@geidea.net>";
                
                // Set subject with proper formatting
                string subjectPrefix = "[POTENTIAL PHISH]";
                if (itemType == "MailItem" && mailItem != null && !string.IsNullOrEmpty(mailItem.Subject))
                {
                    reportEmail.Subject = $"{subjectPrefix} {mailItem.Subject}";
                }
                else
                {
                    reportEmail.Subject = $"{subjectPrefix} {itemType}";
                }

                // Build comprehensive report body
                var reportBody = new StringBuilder();
                reportBody.AppendLine(GetCurrentUserInfos());
                reportBody.AppendLine();
                reportBody.AppendLine(GetBasicInfo(mailItem));
                reportBody.AppendLine();
                reportBody.AppendLine(GetURLsAndAttachmentsInfo(mailItem));
                reportBody.AppendLine();
                reportBody.AppendLine("---------- Email Headers ----------");
                reportBody.AppendLine(headers);
                reportBody.AppendLine();
                reportBody.AppendLine(GetPluginDetails());

                reportEmail.Body = reportBody.ToString();

                // Attach original email
                try
                {
                    reportEmail.Attachments.Add(selectedItem);
                }
                catch
                {
                    // If attachment fails, continue without it
                }

                // Save and send
                reportEmail.Save();
                reportEmail.Send();

                // Update counter
                Properties.Settings.Default.suspicious_reports_counter++;
                Properties.Settings.Default.Save();

                // Show success message
                MessageBox.Show(
                    "✅ Thank you for reporting this suspicious email!\n\n" +
                    "Your report has been sent to the geidea Cybersecurity team for review.\n\n" +
                    "We appreciate your vigilance in helping protect our organization.",
                    "Report Submitted Successfully",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // Delete the reported email
                if (mailItem != null)
                {
                    mailItem.Delete();
                }
            }
            catch (COMException comEx)
            {
                HandleError("Outlook error while creating report", comEx);
            }
            catch (System.Exception ex)
            {
                HandleError("Error processing suspicious email", ex);
            }
        }

        /// <summary>
        /// Handles errors by showing message and sending error report
        /// </summary>
        private void HandleError(string context, System.Exception ex)
        {
            try
            {
                MessageBox.Show(
                    $"An error occurred: {context}\n\n" +
                    "An error report has been automatically sent to support.\n" +
                    "Please contact support if this issue persists.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                // Send error report
                try
                {
                    MailItem errorEmail = Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem) as MailItem;
                    if (errorEmail != null)
                    {
                        errorEmail.To = Properties.Settings.Default.support_email ?? "GRC <grc@geidea.net>";
                        errorEmail.Subject = "[Phishing Reporter Plugin - Error Report]";
                        errorEmail.Body = $"Error Context: {context}\n\n" +
                                        $"Error Message: {ex.Message}\n\n" +
                                        $"Stack Trace:\n{ex.StackTrace}\n\n" +
                                        $"User: {Environment.UserName}\n" +
                                        $"Machine: {Environment.MachineName}\n" +
                                        $"OS: {Environment.OSVersion}\n" +
                                        $"Plugin Version: {Properties.Settings.Default.plugin_version}";
                        errorEmail.Save();
                        errorEmail.Send();
                    }
                }
                catch
                {
                    // If error email fails, at least we tried
                }
            }
            catch
            {
                // Last resort - just show basic error
                MessageBox.Show("A critical error occurred. Please contact support.", "Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Gets basic information about the report
        /// </summary>
        public string GetBasicInfo(MailItem mailItem)
        {
            var info = new StringBuilder();
            info.AppendLine("---------- Basic Information ----------");
            
            try
            {
                if (mailItem != null)
                {
                    var parentFolder = mailItem.Parent as Outlook.MAPIFolder;
                    if (parentFolder != null)
                    {
                        info.AppendLine($" - Reported from: \"{parentFolder.FolderPath}\" Folder");
                    }

                    if (!string.IsNullOrEmpty(mailItem.Subject))
                    {
                        info.AppendLine($" - Email Subject: {mailItem.Subject}");
                    }

                    if (mailItem.ReceivedTime != DateTime.MinValue)
                    {
                        info.AppendLine($" - Received Time: {mailItem.ReceivedTime:yyyy-MM-dd HH:mm:ss}");
                    }
                }
            }
            catch
            {
                info.AppendLine(" - Folder information not available");
            }

            info.AppendLine($" - Operating System: {Environment.OSVersion} {(Environment.Is64BitOperatingSystem ? "(64-bit)" : "(32-bit)")}");
            info.AppendLine($" - Outlook Version: {Globals.ThisAddIn.Application.Name} {Globals.ThisAddIn.Application.Version}");
            info.AppendLine($" - Suspicious emails reported (this session): {Properties.Settings.Default.suspicious_reports_counter}");
            info.AppendLine($" - GoPhish campaigns reported (this session): {Properties.Settings.Default.gophish_reports_counter}");
            info.AppendLine($" - Report Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss UTC}");
            
            return info.ToString();
        }


        /// <summary>
        /// Gets current user information with enhanced error handling
        /// </summary>
        public string GetCurrentUserInfos()
        {
            var userInfo = new StringBuilder();
            userInfo.AppendLine("---------- User Information ----------");
            userInfo.AppendLine($" - Domain: {Environment.UserDomainName ?? "N/A"}");
            userInfo.AppendLine($" - Username: {Environment.UserName ?? "N/A"}");
            userInfo.AppendLine($" - Machine Name: {Environment.MachineName ?? "N/A"}");

            try
            {
                var session = Globals.ThisAddIn.Application.Session;
                if (session?.CurrentUser?.AddressEntry != null)
                {
                    var addrEntry = session.CurrentUser.AddressEntry;
                    if (addrEntry.Type == "EX")
                    {
                        var currentUser = addrEntry.GetExchangeUser();
                        if (currentUser != null)
                        {
                            userInfo.AppendLine($" - Name: {currentUser.Name ?? "N/A"}");
                            userInfo.AppendLine($" - SMTP Address: {currentUser.PrimarySmtpAddress ?? "N/A"}");
                            userInfo.AppendLine($" - Job Title: {currentUser.JobTitle ?? "N/A"}");
                            userInfo.AppendLine($" - Department: {currentUser.Department ?? "N/A"}");
                            userInfo.AppendLine($" - Office Location: {currentUser.OfficeLocation ?? "N/A"}");
                            userInfo.AppendLine($" - Business Phone: {currentUser.BusinessTelephoneNumber ?? "N/A"}");
                            userInfo.AppendLine($" - Mobile Phone: {currentUser.MobileTelephoneNumber ?? "N/A"}");
                        }
                    }
                }
            }
            catch
            {
                userInfo.AppendLine(" - Exchange user information not available");
            }

            return userInfo.ToString();
        }

        /// <summary>
        /// Extracts URLs, domains, and attachment information with enhanced parsing
        /// </summary>
        public string GetURLsAndAttachmentsInfo(MailItem mailItem)
        {
            var info = new StringBuilder();
            info.AppendLine("---------- URLs and Attachments Analysis ----------");

            if (mailItem == null)
            {
                info.AppendLine("Email item not available for analysis");
                return info.ToString();
            }

            var domainsInEmail = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var urlsList = new List<string>();

            // Extract URLs from HTML body
            try
            {
                string emailHTML = mailItem.HTMLBody ?? string.Empty;
                if (!string.IsNullOrEmpty(emailHTML))
                {
                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(emailHTML);

                    var urlNodes = doc.DocumentNode.SelectNodes("//a[@href]");
                    if (urlNodes != null && urlNodes.Count > 0)
                    {
                        info.AppendLine($"\n📎 Total URLs Found: {urlNodes.Count}");

                        foreach (HtmlNode link in urlNodes)
                        {
                            var hrefAttr = link.Attributes["href"];
                            if (hrefAttr != null && !string.IsNullOrWhiteSpace(hrefAttr.Value))
                            {
                                string url = hrefAttr.Value.Trim();
                                urlsList.Add(url);

                                // Sanitize URL for display (replace : with [:] to prevent auto-linking)
                                string sanitizedUrl = url.Replace(":", "[:]");
                                info.AppendLine($"  → URL: {sanitizedUrl}");

                                // Extract domain
                                ExtractDomain(url, domainsInEmail);
                            }
                        }
                    }
                    else
                    {
                        info.AppendLine("\n📎 Total URLs Found: 0");
                    }
                }
                else
                {
                    info.AppendLine("\n📎 HTML body not available for URL extraction");
                }
            }
            catch (System.Exception)
            {
                info.AppendLine("\n⚠️ Error extracting URLs from email");
            }

            // Display unique domains
            if (domainsInEmail.Count > 0)
            {
                info.AppendLine($"\n🌐 Unique Domains Found: {domainsInEmail.Count}");
                foreach (string domain in domainsInEmail.OrderBy(d => d))
                {
                    info.AppendLine($"  → Domain: {domain.Replace(":", "[:]")}");
                }
            }
            else
            {
                info.AppendLine("\n🌐 Unique Domains Found: 0");
            }

            // Process attachments
            try
            {
                info.AppendLine($"\n📎 Total Attachments: {mailItem.Attachments.Count}");

                if (mailItem.Attachments.Count > 0)
                {
                    foreach (Attachment attachment in mailItem.Attachments)
                    {
                        try
                        {
                            string fileName = attachment.FileName ?? "Unknown";
                            long fileSize = attachment.Size;
                            
                            // Generate safe temp file path
                            string safeFileName = Path.GetFileNameWithoutExtension(fileName)
                                .Replace(" ", "_")
                                .Replace("\\", "_")
                                .Replace("/", "_");
                            string tempFilePath = Path.Combine(
                                Path.GetTempPath(),
                                $"Outlook-PhishReporter-{Guid.NewGuid()}-{safeFileName}.tmp");

                            try
                            {
                                attachment.SaveAsFile(tempFilePath);

                                if (File.Exists(tempFilePath))
                                {
                                    string md5Hash = CalculateMD5(tempFilePath);
                                    string sha256Hash = GetHashSha256(tempFilePath);

                                    info.AppendLine($"\n  📎 Attachment: {fileName}");
                                    info.AppendLine($"     Size: {fileSize:N0} bytes");
                                    info.AppendLine($"     MD5: {md5Hash}");
                                    info.AppendLine($"     SHA256: {sha256Hash}");

                                    // Clean up temp file
                                    try { File.Delete(tempFilePath); } catch { }
                                }
                            }
                            catch
                            {
                                info.AppendLine($"\n  📎 Attachment: {fileName} (Size: {fileSize:N0} bytes) - Hash calculation failed");
                            }
                        }
                        catch
                        {
                            info.AppendLine("\n  ⚠️ Error processing attachment");
                        }
                    }
                }
            }
            catch
            {
                info.AppendLine("\n⚠️ Error accessing attachments");
            }

            return info.ToString();
        }

        /// <summary>
        /// Extracts domain from URL with improved error handling
        /// </summary>
        private void ExtractDomain(string url, HashSet<string> domains)
        {
            if (string.IsNullOrWhiteSpace(url))
                return;

            try
            {
                // Handle mailto: links
                if (url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                {
                    int atIndex = url.IndexOf('@');
                    if (atIndex > 0 && atIndex < url.Length - 1)
                    {
                        string email = url.Substring(atIndex + 1);
                        int queryIndex = email.IndexOf('?');
                        if (queryIndex > 0)
                            email = email.Substring(0, queryIndex);
                        
                        if (!string.IsNullOrWhiteSpace(email))
                            domains.Add(email.Trim());
                    }
                    return;
                }

                // Handle regular URLs
                if (Uri.TryCreate(url, UriKind.Absolute, out Uri uri))
                {
                    if (!string.IsNullOrEmpty(uri.Host))
                    {
                        domains.Add(uri.Host);
                    }
                }
                else if (Uri.TryCreate(url, UriKind.RelativeOrAbsolute, out Uri relativeUri))
                {
                    // Try to extract domain from relative URLs that might contain domain info
                    if (url.Contains("@"))
                    {
                        int atIndex = url.IndexOf('@');
                        if (atIndex > 0)
                        {
                            string potentialDomain = url.Substring(atIndex + 1);
                            int slashIndex = potentialDomain.IndexOf('/');
                            if (slashIndex > 0)
                                potentialDomain = potentialDomain.Substring(0, slashIndex);
                            
                            if (!string.IsNullOrWhiteSpace(potentialDomain))
                                domains.Add(potentialDomain.Trim());
                        }
                    }
                }
            }
            catch
            {
                // Silently ignore domain extraction errors
            }
        }



        /// <summary>
        /// Gets plugin details and version information
        /// </summary>
        public string GetPluginDetails()
        {
            var details = new StringBuilder();
            details.AppendLine("---------- Phishing Reporter Plugin Details ----------");
            details.AppendLine($" - Version: {Properties.Settings.Default.plugin_version ?? "Unknown"}");
            details.AppendLine($" - Purpose: Report phishing emails to geidea Cybersecurity team");
            details.AppendLine($" - Support Email: {Properties.Settings.Default.support_email ?? "N/A"}");
            details.AppendLine($" - Report Generated: {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC");
            return details.ToString();
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PhishingReporter.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Calculates MD5 hash of a file
        /// </summary>
        private static string CalculateMD5(string filename)
        {
            if (string.IsNullOrEmpty(filename) || !File.Exists(filename))
                return "N/A";

            try
            {
                using (var md5 = MD5.Create())
                {
                    using (var stream = File.OpenRead(filename))
                    {
                        byte[] hash = md5.ComputeHash(stream);
                        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                    }
                }
            }
            catch
            {
                return "Error calculating MD5";
            }
        }

        /// <summary>
        /// Calculates SHA256 hash of a file
        /// </summary>
        private string GetHashSha256(string filename)
        {
            if (string.IsNullOrEmpty(filename) || !File.Exists(filename))
                return "N/A";

            try
            {
                using (FileStream stream = File.OpenRead(filename))
                {
                    using (SHA256 sha = SHA256.Create())
                    {
                        byte[] shaHash = sha.ComputeHash(stream);
                        var result = new StringBuilder(64);
                        foreach (byte b in shaHash)
                        {
                            result.Append(b.ToString("x2"));
                        }
                        return result.ToString();
                    }
                }
            }
            catch
            {
                return "Error calculating SHA256";
            }
        }

        #endregion
    }

    public static class MailItemExtensions
    {
        private const string HeaderRegex =
            @"^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)" +
                "(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)";
        private const string TransportMessageHeadersSchema =
            "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        public static string[] Headers(this MailItem mailItem, string name)
        {
            var headers = mailItem.HeaderLookup();
            if (headers.Contains(name))
                return headers[name].ToArray();
            return new string[0];
        }

        public static ILookup<string, string> HeaderLookup(this MailItem mailItem)
        {
            var headerString = mailItem.HeaderString();
            var headerMatches = Regex.Matches
                (headerString, HeaderRegex, RegexOptions.Multiline).Cast<Match>();
            return headerMatches.ToLookup(
                h => h.Groups["header_key"].Value,
                h => h.Groups["header_value"].Value);
        }

        public static string HeaderString(this MailItem mailItem)
        {
            return (string)mailItem.PropertyAccessor
                .GetProperty(TransportMessageHeadersSchema);
        }

    }
}
 