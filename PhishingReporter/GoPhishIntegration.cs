/*
 * Developer: Abdulla Albreiki
 * Github: https://github.com/0dteam
 * Licensed under the GNU General Public License v3.0
 * Enhanced for production - Modern C# implementation
 */

using System;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Security.Authentication;
using System.Text;

namespace PhishingReporter
{
    /// <summary>
    /// Modern GoPhish integration with improved error handling and security
    /// </summary>
    static class GoPhishIntegration
    {
        private static readonly HttpClient httpClient;
        private static readonly string GoPhishHeader;
        private static readonly Regex WebExpID;
        private static readonly string WebExpPrefix;

        static GoPhishIntegration()
        {
            // Configure TLS support - flexible for internal networks
            // Supports both HTTP (internal) and HTTPS (external) connections
            try
            {
                // Try to enable TLS 1.2/1.3 for HTTPS connections
                // But allow HTTP connections for internal networks without TLS
                ServicePointManager.SecurityProtocol = 
                    SecurityProtocolType.Tls12 | 
                    SecurityProtocolType.Tls13 | 
                    SecurityProtocolType.Ssl3 | 
                    SecurityProtocolType.Tls;
            }
            catch
            {
                // Fallback: Allow all protocols including HTTP for internal networks
                try
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls;
                }
                catch
                {
                    // If TLS configuration fails, continue without forcing TLS
                    // This allows HTTP connections in internal networks
                }
            }

            // Initialize HttpClient - supports both HTTP and HTTPS
            var handler = new HttpClientHandler
            {
                AllowAutoRedirect = true,
                MaxAutomaticRedirections = 3,
                // Don't validate SSL certificates for internal networks
                ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
            };
            
            httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(10) // 10 second timeout
            };
            httpClient.DefaultRequestHeaders.Add("User-Agent", "PhishingReporter-Plugin/1.1");

            // Initialize GoPhish settings
            GoPhishHeader = Properties.Settings.Default.gophish_custom_header ?? "X-GOPHISH-AJSMN";
            WebExpID = new Regex(GoPhishHeader + @":\s*([0-9a-zA-Z]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
            WebExpPrefix = GoPhishHeader + @": ";
        }

        /// <summary>
        /// Constructs GoPhish report URL from custom header in simulated phishing campaign email
        /// </summary>
        /// <param name="headers">Email headers string</param>
        /// <returns>Report URL or null if header not found</returns>
        public static string SetReportURL(string headers)
        {
            if (string.IsNullOrWhiteSpace(headers))
                return null;

            try
            {
                // Extract GoPhish Custom Header (X-GOPHISH-AJSMN: USERID0123)
                var match = WebExpID.Match(headers);
                
                if (match.Success && match.Groups.Count > 1)
                {
                    // Extract User ID from the header (USERID0123)
                    string userId = match.Groups[1].Value.Trim();
                    
                    if (!string.IsNullOrEmpty(userId))
                    {
                        // Build reporting URL: http://GOPHISHURL:PORT/report?rid=USERID
                        string goPhishUrl = Properties.Settings.Default.gophish_url ?? "http://10.100.125.230";
                        string port = Properties.Settings.Default.gophish_listener_port ?? "3333";
                        string reportUrl = $"{goPhishUrl}:{port}/report?rid={Uri.EscapeDataString(userId)}";
                        return reportUrl;
                    }
                }
            }
            catch (System.Exception)
            {
                // Log error silently, return null to indicate failure
            }

            return null;
        }

        /// <summary>
        /// Sends report notification to GoPhish server asynchronously
        /// </summary>
        /// <param name="reportURL">The GoPhish report URL</param>
        /// <returns>True if successful, false otherwise</returns>
        public static bool SendReportNotificationToServer(string reportURL)
        {
            if (string.IsNullOrWhiteSpace(reportURL))
                return false;

            try
            {
                // Validate URL format
                if (!Uri.TryCreate(reportURL, UriKind.Absolute, out Uri validatedUri))
                    return false;

                // Use synchronous call for VSTO compatibility (async not well supported in Office add-ins)
                var response = httpClient.GetAsync(validatedUri).GetAwaiter().GetResult();
                
                if (response.IsSuccessStatusCode)
                {
                    return true;
                }
            }
            catch (HttpRequestException)
            {
                // Network error - GoPhish server not reachable
            }
            catch (TaskCanceledException)
            {
                // Timeout occurred
            }
            catch (System.Exception)
            {
                // Other errors
            }

            return false;
        }

        /// <summary>
        /// Cleanup resources
        /// </summary>
        public static void Dispose()
        {
            httpClient?.Dispose();
        }
    }
}

