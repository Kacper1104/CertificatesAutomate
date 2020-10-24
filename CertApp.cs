using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CertApp
{
    class CertApp
    {
        string templatePath;
        string templateDocumentText;
        public CertApp()
        {
            templatePath = Path.Combine(Environment.CurrentDirectory, @GlobalVariables.RESOURCES_LOCATION, GlobalVariables.TEMPLATE_FILE_NAME);

            WordprocessingDocument template = OpenFile(templatePath);

            StreamReader sr = new StreamReader(template.MainDocumentPart.GetStream());
            templateDocumentText = sr.ReadToEnd();
            sr.Close();
            template.Close();

        }
        private WordprocessingDocument OpenFile(string filePath)
        {
            return WordprocessingDocument.Open(
                filePath,
                true);
        }
        public void WriteNewFile(string destinationDirectory, string destinationFileName, string textToReplace)
        {
            string destinationFilePath = Path.Combine(Environment.CurrentDirectory, @GlobalVariables.DESTINATION_LOCATION, destinationDirectory + "\\" + destinationFileName);
            Regex regexText = new Regex(GlobalVariables.TEXT_PLACEHOLDER);
            string newDocumentText = regexText.Replace(templateDocumentText, textToReplace);
            try
            {
                File.Copy(templatePath, destinationFilePath);
            }
            catch (System.IO.DirectoryNotFoundException exception)
            {
                Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory, @GlobalVariables.DESTINATION_LOCATION, destinationDirectory));
                File.Copy(templatePath, destinationFilePath);
            }

            WordprocessingDocument destinationFile = OpenFile(destinationFilePath);
            StreamWriter sw = new StreamWriter(destinationFile.MainDocumentPart.GetStream(FileMode.Create));
            sw.Write(newDocumentText);
            sw.Close();
            destinationFile.Close();

        }
    }
}
