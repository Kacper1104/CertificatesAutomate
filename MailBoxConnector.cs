using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Net.Mail;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;

namespace CertApp
{
    class MailBoxConnector
    {
        SmtpClient smtp;
        MailAddress fromAddress;
        string fromPassword;
        WordJPGConverter WJC;
        public MailBoxConnector(WordJPGConverter WJC)
        {
            this.WJC = WJC;
            fromAddress = new MailAddress(GlobalVariables.SENDER_EMAIL, GlobalVariables.SENDER_NAME);
            fromPassword = GlobalVariables.SENDER_PASSWORD;
            smtp = new SmtpClient()
            {
                Host = GlobalVariables.SMTP_ADDRESS,
                Port = GlobalVariables.SMTP_PORT,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
            };
        }

        public void SendMail(string filedir, bool isProd, string batch)
        {
            var receivers = filedir;
            string[] to;
            if (isProd)
            {
                to = receivers.Split(',');
            }
            else
            {
                to = GlobalVariables.SENDER_EMAIL.Split(',');
            }
            const string subject = GlobalVariables.MSG_SUBJECT;
            string template = File.ReadAllText(Path.Combine(Environment.CurrentDirectory, @GlobalVariables.RESOURCES_LOCATION, "email.html"));
            string body = template;
            var message = new MailMessage();
            message.Subject = subject;
            message.IsBodyHtml = true;
            message.Body = body;
            for (int i = 0; i < to.Length; i++)
            {
                message.To.Add(new MailAddress(to[i].Trim(), to[i].Trim()));
            }
            message.From = fromAddress;

            //attachments
            string path = Path.Combine(Environment.CurrentDirectory, @GlobalVariables.DESTINATION_LOCATION, batch + "\\" + filedir);
            string[] files = Directory.GetFiles(path);
            for (int i = 0; i < files.Length; i++)
            {
                if (files[i].EndsWith("docx")) continue;
                Console.WriteLine(files[i]);
                try
                {
                    //WJC.Print(files[i]);
                }
                catch (Exception e)
                { }
                message.Attachments.Add(new Attachment(files[i]));
            }
            smtp.Send(message);
        }
    }


}
