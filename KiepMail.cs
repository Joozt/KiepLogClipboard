using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Prabhu;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Windows;

namespace KiepLogClipboard
{
    class KiepMail
    {
        private const string hostname = "pop.gmail.com";
        private const int port = 995;


        // TODO Insert Gmail account details
        private const string emailAddress = "my-email@gmail.com";
        private const string password = "my-password";


        private static System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("nl-NL");
        private static Regex CharsetRegex = new Regex("charset=\"?(?<charset>[^\\s\"]+)\"?", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static Regex QuotedPrintableRegex = new Regex("=(?<hexchars>[0-9a-fA-F]{2,2})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static Regex UrlRegex = new Regex("(?<url>https?://[^\\s\"]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static Regex FilenameRegex = new Regex("filename=\"?(?<filename>[^\\s\"]+)\"?", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static Regex NameRegex = new Regex("name=\"?(?<filename>[^\\s\"]+)\"?", RegexOptions.IgnoreCase | RegexOptions.Compiled);


        public string Download(DateTime beginDate, DateTime endDate)
        {
            string result = String.Empty;

            if (endDate < beginDate)
            {
                return "End date before begin date";
            }

            using (Prabhu.Pop3Client client = new Prabhu.Pop3Client(hostname, port, emailAddress, password, true))
            {
                client.Connect();
                Email email;
                Email nextEmail;
                DateTime dateEmail;
                DateTime nextDateEmail;

                Console.WriteLine("Number of e-mails: " + client.GetEmailCount());

                // Loop from last to first e-mail
                for (int i = client.GetEmailCount(); i > 0; i--)
                {
                    // Get e-mail
                    email = client.FetchEmail(i);

                    // Take e-mail date
                    dateEmail = email.UtcDateTime.Date;

                    // Init next date to current e-mail date
                    nextDateEmail = dateEmail;

                    // Get the next e-mail
                    if (i > 1)
                    {
                        // Get next e-mail
                        nextEmail = client.FetchEmail(i - 1);

                        // Take next e-mail date
                        nextDateEmail = nextEmail.UtcDateTime.Date;
                    }

                    // Do nothing until within date range
                    if (dateEmail > endDate)
                    {
                        Console.WriteLine("dateEmail > Einddatum: " + dateEmail + " > " + endDate);
                        continue;
                    }

                    // Stop processing if date is before begin date
                    else if (dateEmail < beginDate)
                    {
                        Console.WriteLine("Emaildatum < begindatum: " + dateEmail + " < " + beginDate);
                        break;
                    }

                    // Process e-mails within date range
                    else if (dateEmail >= beginDate && dateEmail <= endDate)
                    {

                        // Get content (MessagePart is a line of the e-mail)
                        List<MessagePart> msgParts = client.FetchMessageParts(i);
                        if (email != null && msgParts != null)
                        {
                            MessagePart preferredMsgPart = FindMessagePart(msgParts, "text/html");
                            if (preferredMsgPart == null)
                                preferredMsgPart = FindMessagePart(msgParts, "text/plain");
                            else if (preferredMsgPart == null && msgParts.Count > 0)
                                preferredMsgPart = msgParts[0];
                            string contentType, charset, contentTransferEncoding, body = null;
                            if (preferredMsgPart != null)
                            {
                                contentType = preferredMsgPart.Headers["Content-Type"];
                                charset = "us-ascii";
                                contentTransferEncoding = preferredMsgPart.Headers["Content-Transfer-Encoding"];
                                Match m = CharsetRegex.Match(contentType);
                                if (m.Success)
                                    charset = m.Groups["charset"].Value;
                                if (contentTransferEncoding != null)
                                {
                                    if (contentTransferEncoding.ToLower() == "base64")
                                        body = DecodeBase64String(charset, preferredMsgPart.MessageText);
                                    else if (contentTransferEncoding.ToLower() == "quoted-printable")
                                        body = DecodeQuotedPrintableString(preferredMsgPart.MessageText);
                                    else
                                        body = preferredMsgPart.MessageText;
                                }
                                else
                                    body = preferredMsgPart.MessageText;
                            }
                            result = (preferredMsgPart != null ? (preferredMsgPart.Headers["Content-Type"].IndexOf("text/plain") != -1 ? FormatUrls(body) + "\n\n" : body) : null) + result;

                            if (dateEmail > nextDateEmail || i == 1)
                            {
                                result = dateEmail.ToString("dddd d MMMM yyyy", ci) + "\n" + result;
                            }
                        }
                    }
                    else
                    {
                        result = "Unexpected error!" + result;
                    }

                }

                if (result == String.Empty)
                {
                    return "No e-mails found!";
                }
                else
                {
                    result = result.Remove(0, 4);
                    return result;
                }

            }
        }

        protected Decoder GetDecoder(string charset)
        {
            Decoder decoder;
            switch (charset.ToLower())
            {
                case "utf-7":
                    decoder = Encoding.UTF7.GetDecoder();
                    break;
                case "utf-8":
                    decoder = Encoding.UTF8.GetDecoder();
                    break;
                case "us-ascii":
                    decoder = Encoding.ASCII.GetDecoder();
                    break;
                case "iso-8859-1":
                    decoder = Encoding.ASCII.GetDecoder();
                    break;
                default:
                    decoder = Encoding.ASCII.GetDecoder();
                    break;
            }
            return decoder;
        }

        protected string DecodeBase64String(string charset, string encodedString)
        {
            Decoder decoder = GetDecoder(charset);
            byte[] buffer = Convert.FromBase64String(encodedString);
            char[] chararr = new char[decoder.GetCharCount(buffer, 0, buffer.Length)];
            decoder.GetChars(buffer, 0, buffer.Length, chararr, 0);
            return new string(chararr);
        }

        protected string DecodeQuotedPrintableString(string encodedString)
        {
            StringBuilder b = new StringBuilder();
            int startIndx = 0;
            MatchCollection matches = QuotedPrintableRegex.Matches(encodedString);
            for (int i = 0; i < matches.Count; i++)
            {
                Match m = matches[i];
                string hexchars = m.Groups["hexchars"].Value;
                int charcode = Convert.ToInt32(hexchars, 16);
                char c = (char)charcode;
                if (m.Index > 0)
                    b.Append(encodedString.Substring(startIndx, (m.Index - startIndx)));
                b.Append(c);
                startIndx = m.Index + 3;
            }
            if (startIndx < encodedString.Length)
                b.Append(encodedString.Substring(startIndx));
            return Regex.Replace(b.ToString(), "=\r\n", "");
        }

        protected MessagePart FindMessagePart(List<MessagePart> msgParts, string contentType)
        {
            foreach (MessagePart p in msgParts)
                if (p.ContentType != null && p.ContentType.IndexOf(contentType) != -1)
                    return p;
            return null;
        }

        protected string FormatUrls(string plainText)
        {
            string replacementLink = "<a href=\"${url}\">${url}</a>";
            return UrlRegex.Replace(plainText, replacementLink);
        }
    }
}
