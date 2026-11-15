using System;
using System.Security.Cryptography;
using System.Text;

namespace JameelCopilot.Services
{
    /// <summary>
    /// The Custom API key as stored in the registry: encrypted with Windows
    /// DPAPI for this machine, base64. The installer writes the same format.
    /// This keeps the key out of plain text on disk and in backups; it does
    /// not hide it from a user of this PC, so the key must be a restricted one.
    /// </summary>
    public static class ApiKeyProtector
    {
        public static string Protect(string apiKey)
        {
            if (string.IsNullOrEmpty(apiKey)) throw new ArgumentException("The API key is empty.", nameof(apiKey));
            return Convert.ToBase64String(
                ProtectedData.Protect(Encoding.UTF8.GetBytes(apiKey), null, DataProtectionScope.LocalMachine));
        }

        /// <summary>The key, or null when the value is blank, not base64, or not decryptable on this PC.</summary>
        public static string Unprotect(string protectedBase64)
        {
            if (string.IsNullOrWhiteSpace(protectedBase64)) return null;
            try
            {
                var key = Encoding.UTF8.GetString(ProtectedData.Unprotect(
                    Convert.FromBase64String(protectedBase64.Trim()), null, DataProtectionScope.LocalMachine));
                return key.Length == 0 ? null : key;
            }
            catch (Exception ex) when (ex is FormatException || ex is CryptographicException)
            {
                return null;
            }
        }
    }
}
