using System;
using System.Collections.Generic;
using System.Text;

namespace CertApp
{
    public static class GlobalVariables
    {
        public const string RESOURCES_LOCATION = "Resources\\";
        public const string DESTINATION_LOCATION = "Results\\";
        public const string TEMPLATE_FILE_NAME = "TEMPLATE.docx";
        public const string REPORT_FILE_NAME = "SOURCE.csv";
        public const string TEXT_PLACEHOLDER = "«NAME AND SURNAME»";
        public const string SENDER_EMAIL = "email@email.com";
        public const string SENDER_NAME = "NAME SURNAME";
        public const string SENDER_PASSWORD = "Password123";
        public const string PDF_PRINTER = "Microsoft Print to PDF";
        public const string MSG_SUBJECT = "Message subject";
        public const string SMTP_ADDRESS = "smtp.gmail.com";
        public const int SMTP_PORT = 587;
    }
}
