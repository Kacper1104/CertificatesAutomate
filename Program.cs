using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Threading;
using System.Windows.Documents;
using System.Windows.Xps.Packaging;

namespace CertApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //    CertApp app = new CertApp();
            //    Parser parser = new Parser();
            //    List<string> line;
            //    while ((line = parser.ReadLine()) != null)
            //    {
            //        for (int i = 1; i < line.Count; i++)
            //        {
            //            try
            //            {
            //                app.WriteNewFile(line[0], line[i] + ".docx", line[i]);
            //            }
            //            catch (Exception e)
            //            {
            //                //Console.WriteLine(e.ToString());
            //            }
            //        }
            //    }
            //    parser.PrintLongest();

            //Generate pdfs and send emails
            //WordJPGConverter WJC = new WordJPGConverter();
            //MailBoxConnector MBC = new MailBoxConnector(WJC);
            //string[] batch = new string[] { "batch1", "batch2", "batch3", "batch4", "batch5", "batch6", "batch7", "batch8", "batch9", "batch10", "batch11", "batch12", "batch13", "batch14", "batch15", "batch16" };
            //for (int j = 3; j < 16; j++)
            //{
            //    string path = Path.Combine(Environment.CurrentDirectory, GlobalVariables.DESTINATION_LOCATION + batch[j] + "\\");
            //    string[] directories = Directory.GetDirectories(path);
            //    for (int i = 0; i < directories.Length; i++)
            //    {
            //        var dir = new DirectoryInfo(directories[i]);
            //        MBC.SendMail(dir.Name, true, batch[j]);
            //        Thread.Sleep(TimeSpan.FromSeconds(1));
            //    }
            //}
        }
    }
}
