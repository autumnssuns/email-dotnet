using System.Net.Mail;
using System.Configuration;
using System.Collections.Specialized;
using Microsoft.Graph.Models.ExternalConnectors;
using System.Runtime.CompilerServices;

namespace EmailDotnet
{
    class EmailClient
    {
        private static string host;
        private static int port;
        private static bool ssl;

        private SmtpClient client = new SmtpClient(host);
        private string email;

        public static void ConfigHost(string host, int port, bool ssl)
        {
            EmailClient.host = host;
            EmailClient.port = port;
            EmailClient.ssl = ssl;
        }

        public EmailClient(string email, string password)
        {
            this.email = email;
            client.Port = port;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential(email, password);
            client.EnableSsl = ssl;
        }

        public void Send(Email email)
        {
            MailMessage message = new MailMessage(this.email, email.To);
            message.Subject = email.Subject;
            message.Body = email.Body;
            foreach (Attachment attachment in GetAttachments(email.AttachmentPath))
            {
                message.Attachments.Add(attachment);
            }
            client.Send(message);
        }

        private IEnumerable<Attachment> GetAttachments(string path)
        {
            if (File.Exists(path))
            {
                yield return new Attachment(path);
            }
            else
            {
                foreach (string file in Directory.EnumerateFiles(path))
                {
                    yield return new Attachment(file);
                }
            }
        }
    }
}
