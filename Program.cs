using ExcelDataReader;
using Microsoft.Extensions.Configuration;
using EmailDotnet;
using System.Net.Mail;
using System.Net.Mime;
class Program
{
    static void Main(string[] args)
    {
        var configBuilder = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
        var configSection = configBuilder.GetSection("Smtp");
        var host = configSection["host"] ?? null;
        var port = configSection["port"] ?? null;
        var ssl = configSection["ssl"] ?? null;

        try
        {
            if (host == null || port == null || ssl == null)
            {
                throw new MissingFieldException("Missing host, port or ssl in appsettings.json");
            }
            EmailClient.ConfigHost(host, int.Parse(port), bool.Parse(ssl));
            
            var email = args[0];
            var password = args[1];
            IEnumerable<Email> emails = GetTargets();
            int count = 0;
            foreach (Email mail in emails)
            {
                Console.WriteLine(mail);
                count++;
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("{0} emails found. Enter Y to send or N to cancel: ", count);
            string? confirm = Console.ReadLine()?.ToUpper();
            while (confirm != "Y" && confirm != "N")
            {
                Console.Write("Invalid input. Enter Y to send or N to cancel: ");
                confirm = Console.ReadLine()?.ToUpper();
            }
            if (confirm == "N") return;
            Console.ResetColor();

            EmailClient client = new EmailClient(email, password);
            int current = 1;
            foreach (Email mail in emails)
            {
                Console.Write("Sending email {0} of {1}... ", current, count);
                client.Send(mail);
                Console.WriteLine("Done!");
                current++;
            }
        }
        catch (IndexOutOfRangeException)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Missing email address or password");
        }
        catch (MissingFieldException)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Missing host or port in appsettings.json");
        }
        catch (Exception e)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("An unexpected exception occured: ");
            Console.WriteLine(e.Message);
        }
        finally
        {
            Console.ResetColor();
        }
    }

    static IEnumerable<Email> GetTargets()
    {
        string PATH = "Targets.xlsx";
        LinkedList<Email> emails = new LinkedList<Email>();
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using (var stream = File.Open(PATH, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    reader.Read(); // skip header row
                    while (reader.Read())
                    {
                        if (reader.GetString(0) == null) continue;
                        emails.AddLast(new Email
                        {
                            To = reader.GetString(0),
                            AttachmentPath = reader.GetString(1),
                            Subject = reader.GetString(2),
                            Body = reader.GetString(3)
                        });
                    }
                } while (reader.NextResult());
            }
        }
        return emails;
    }
}