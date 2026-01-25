#region License statement
/* SnakeTail is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, version 3 of the License.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
#endregion

using System;
using System.IO;
using System.Text;

namespace SnakeTail
{
    /// <summary>
    /// Helper class for encoding detection and handling
    /// </summary>
    public static class EncodingHelper
    {
        /// <summary>
        /// Detects the encoding of a file, with special handling for UTF8 and UTF8 BOM
        /// </summary>
        /// <param name="filePath">Path to the file to detect</param>
        /// <returns>Detected encoding, or Encoding.Default if detection fails</returns>
        public static Encoding DetectFileEncoding(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return Encoding.Default;

            try
            {
                // Read first 4 bytes to check for BOM
                byte[] bom = new byte[4];
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    int bytesRead = fs.Read(bom, 0, 4);
                    if (bytesRead < 3)
                        return Encoding.Default;

                    // Check for UTF8 BOM (EF BB BF)
                    if (bytesRead >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
                    {
                        return new UTF8Encoding(true); // UTF8 with BOM
                    }

                    // Check for UTF16 LE BOM (FF FE)
                    if (bytesRead >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
                    {
                        return Encoding.Unicode;
                    }

                    // Check for UTF16 BE BOM (FE FF)
                    if (bytesRead >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
                    {
                        return Encoding.BigEndianUnicode;
                    }

                    // Try to detect UTF8 without BOM by reading a sample
                    fs.Seek(0, SeekOrigin.Begin);
                    byte[] buffer = new byte[Math.Min(8192, fs.Length)];
                    bytesRead = fs.Read(buffer, 0, buffer.Length);

                    if (bytesRead > 0)
                    {
                        // Try to decode as UTF8 and see if it's valid
                        try
                        {
                            UTF8Encoding utf8NoBom = new UTF8Encoding(false);
                            string test = utf8NoBom.GetString(buffer, 0, bytesRead);
                            // If decoding succeeds, check if it contains valid UTF8 sequences
                            // and if it has non-ASCII characters
                            bool hasNonAscii = false;
                            for (int i = 0; i < test.Length; i++)
                            {
                                if (test[i] > 127)
                                {
                                    hasNonAscii = true;
                                    break;
                                }
                            }
                            // If we have non-ASCII characters or the byte pattern looks like UTF8
                            if (hasNonAscii || IsLikelyUtf8(buffer, bytesRead))
                            {
                                return new UTF8Encoding(false);
                            }
                        }
                        catch
                        {
                            // Decoding failed, try pattern matching
                            if (IsLikelyUtf8(buffer, bytesRead))
                            {
                                return new UTF8Encoding(false);
                            }
                        }
                    }
                }
            }
            catch
            {
                // If detection fails, return default
            }

            return Encoding.Default;
        }

        /// <summary>
        /// Checks if a byte array is likely UTF8 encoding (more lenient check)
        /// </summary>
        private static bool IsLikelyUtf8(byte[] buffer, int length)
        {
            if (length == 0)
                return false;

            int validUtf8Bytes = 0;
            int totalNonAsciiBytes = 0;
            int i = 0;

            while (i < length)
            {
                byte b = buffer[i];
                if (b < 0x80)
                {
                    // ASCII character
                    i++;
                }
                else
                {
                    totalNonAsciiBytes++;

                    if ((b & 0xE0) == 0xC0 && i + 1 < length && (buffer[i + 1] & 0xC0) == 0x80)
                    {
                        // Valid 2-byte UTF8 character
                        validUtf8Bytes += 2;
                        i += 2;
                    }
                    else if ((b & 0xF0) == 0xE0 && i + 2 < length &&
                             (buffer[i + 1] & 0xC0) == 0x80 && (buffer[i + 2] & 0xC0) == 0x80)
                    {
                        // Valid 3-byte UTF8 character
                        validUtf8Bytes += 3;
                        i += 3;
                    }
                    else if ((b & 0xF8) == 0xF0 && i + 3 < length &&
                             (buffer[i + 1] & 0xC0) == 0x80 && (buffer[i + 2] & 0xC0) == 0x80 &&
                             (buffer[i + 3] & 0xC0) == 0x80)
                    {
                        // Valid 4-byte UTF8 character
                        validUtf8Bytes += 4;
                        i += 4;
                    }
                    else
                    {
                        // Invalid UTF8 sequence, skip one byte
                        i++;
                    }
                }
            }

            // If we have non-ASCII bytes and most of them are valid UTF8, consider it UTF8
            // Require at least 50% of non-ASCII bytes to be valid UTF8 sequences
            if (totalNonAsciiBytes > 0)
            {
                return (validUtf8Bytes * 100 / totalNonAsciiBytes) >= 50;
            }

            // If all ASCII, could be UTF8 or ASCII, but prefer UTF8 for detection
            return true;
        }

        /// <summary>
        /// Gets a display name for an encoding
        /// </summary>
        public static string GetEncodingDisplayName(Encoding encoding)
        {
            if (encoding == null)
                return "Default";

            if (encoding is UTF8Encoding)
            {
                UTF8Encoding utf8 = (UTF8Encoding)encoding;
                // Check if it's UTF8 with BOM by creating a new instance and comparing
                UTF8Encoding utf8WithBom = new UTF8Encoding(true);
                UTF8Encoding utf8WithoutBom = new UTF8Encoding(false);

                // UTF8Encoding doesn't expose BOM info directly, so we check by encoding a test string
                // But a simpler approach: check the preamble
                if (encoding.GetPreamble().Length > 0)
                    return "UTF8 BOM";
                else
                    return "UTF8";
            }
            else if (encoding == Encoding.ASCII)
            {
                return "ASCII";
            }
            else if (encoding == Encoding.Unicode)
            {
                return "Unicode (UTF16 LE)";
            }
            else if (encoding == Encoding.Default)
            {
                return "Default";
            }
            else
            {
                return encoding.EncodingName;
            }
        }

        /// <summary>
        /// Creates an encoding from a display name
        /// </summary>
        public static Encoding GetEncodingFromDisplayName(string displayName)
        {
            if (string.IsNullOrEmpty(displayName))
                return Encoding.Default;

            switch (displayName.ToUpper())
            {
                case "UTF8":
                case "UTF-8":
                    return new UTF8Encoding(false);
                case "UTF8 BOM":
                case "UTF-8 BOM":
                    return new UTF8Encoding(true);
                case "ASCII":
                    return Encoding.ASCII;
                case "UNICODE":
                case "UNICODE (UTF16 LE)":
                case "UTF16":
                case "UTF-16":
                    return Encoding.Unicode;
                case "DEFAULT":
                default:
                    return Encoding.Default;
            }
        }
    }
}
