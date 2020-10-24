using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace CertApp
{
    class Parser
    {
        StreamReader sr;
        string reportFilePath;
        List<string> lines;
        int currentLine = 0;
        int longestName = 75;
        List<string> longestNames;
        List<string> longestEmails;
        public Parser()
        {
            longestNames = new List<string>();
            longestEmails = new List<string>();
            reportFilePath = Path.Combine(Environment.CurrentDirectory, @GlobalVariables.RESOURCES_LOCATION, GlobalVariables.REPORT_FILE_NAME);
            sr = new StreamReader(reportFilePath);
            string line;
            lines = new List<string>();
            while ((line = sr.ReadLine()) != null)
            {
                lines.Add(Regex.Replace(line, @";+", "|"));
            }
        }
        public List<string> ReadLine()
        {
            List<string> elements = new List<string>();
            bool emailNoted = false;
            string line = null;
            try
            {
                line = lines[currentLine];
            }
            catch (Exception e)
            {
                //Console.WriteLine(e.ToString());
                return null;
            }
            string[] split = line.Split('|');
            foreach (string element in split)
            {
                if (element.Trim() != "")
                {
                    elements.Add(element.Trim());
                    if (element.Trim().Length > longestName)
                    {
                        longestNames.Add(element.Trim());
                        if (!emailNoted)
                        {
                            emailNoted = true;
                            longestEmails.Add(split[0]);
                        }
                    }
                }
            }

            currentLine++;
            return elements;
        }
        public void PrintLongest()
        {
            Console.WriteLine("Long elements: " + longestNames.Count);
            Console.WriteLine();
            for (int i = 0; i < longestNames.Count; i++)
            {
                Console.WriteLine(longestNames[i]);
            }
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("Long elements directories: ");
            Console.WriteLine();
            for (int i = 0; i < longestEmails.Count; i++)
            {
                Console.WriteLine(longestEmails[i]);
            }
        }
    }
}
