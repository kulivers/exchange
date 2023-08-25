using System;
using System.Net;
using System.Net.Security;
using Microsoft.Exchange.WebServices.Data;

namespace exchange
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            var credential = new NetworkCredential("ekul", "Xotourlife9991!", "corp");
            var ass =  new ExchangeService(ExchangeVersion.Exchange2013,
                TimeZoneInfo.CreateCustomTimeZone("Часовой пояс",
                    TimeZoneInfo.Local.GetUtcOffset(DateTime.Now),
                    "Часовой пояс",
                    "Часовой пояс"))
            {
                Timeout = 100000,
                Credentials = credential,
                Url = new Uri("https://mail.comindware.ru/EWS/Exchange.asmx"),
            };
            try
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) =>
                {
                    var month = ((DateTime)certificate.GetType().GetProperty("NotAfter").GetValue(certificate)).Month;
                    Console.WriteLine("certificate: " + certificate.Issuer + " " + certificate.GetExpirationDateString() + " " + sslPolicyErrors);
                    Console.WriteLine($"(month = {month})");
                    if (sslPolicyErrors!=SslPolicyErrors.None)
                    {
                        return false;
                    }
                    return true;
                };
                ass.FindItems(WellKnownFolderName.Inbox, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false),
                    new ItemView(1));
            }
            catch (Exception e)
            {
                while (e != null)
                {
                    e = e.InnerException;
                }
            }
        }
    }
}