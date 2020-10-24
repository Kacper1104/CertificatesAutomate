using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Text;
using System.Windows.Controls;

namespace CertApp
{
    class WordJPGConverter
    {

        public void Convert(string filepathandname)
        {
            Application myWordApp = new Application();
            Document myWordDoc = new Document();
            object missing = System.Type.Missing;
            string path1 = filepathandname;
            myWordDoc = myWordApp.Documents.Add(path1, missing, missing, missing);

            foreach (Microsoft.Office.Interop.Word.Window window in myWordDoc.Windows)
            {
                foreach (Microsoft.Office.Interop.Word.Pane pane in window.Panes)
                {
                    for (var i = 1; i <= pane.Pages.Count; i++)
                    {
                        var bits = pane.Pages[i].EnhMetaFileBits;
                        try
                        {
                            using (var ms = new MemoryStream((byte[])(bits)))
                            {
                                var image = System.Drawing.Image.FromStream(ms);
                                var jpegTarget = Path.ChangeExtension(path1, "jpeg");
                                image.Save(jpegTarget, System.Drawing.Imaging.ImageFormat.Jpeg);
                            }
                        }
                        catch (System.Exception ex)
                        { }
                    }
                }
            }
        }

        public void Print(string filepathandname)
        {
            WdSaveFormat format = WdSaveFormat.wdFormatPDF;
            // Create an instance of Word.exe
            _Application oWord = new Application
            {

                // Make this instance of word invisible (Can still see it in the taskmgr).
                Visible = false
            };

            // Interop requires objects.
            object oMissing = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = true;     // Does not cause any word dialog to show up
                                        //object readOnly = false;  // Causes a word object dialog to show at the end of the conversion
            object oInput = filepathandname;
            object oOutput = filepathandname.Substring(0, filepathandname.Length - 4) + "pdf";
            object oFormat = format;
            object oPrintToFile = true;

            // Load a document into our instance of word.exe
            _Document oDoc = oWord.Documents.Open(
                ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                );

            // Make this document the active document.
            oDoc.Activate();

            // Save this document using Word
            //oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing,
            //    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            //    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
            //    );

            oWord.Visible = false;

            PrintDialog pDialog = new PrintDialog();
            oWord.ActivePrinter = GlobalVariables.PDF_PRINTER;
            oDoc.PrintOut(ref oMissing, ref oMissing, ref oMissing, ref oOutput,
                 ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                 ref oMissing, ref oMissing, ref oPrintToFile, ref oMissing,
                 ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                 ref oMissing, ref oMissing);
            oDoc.Close(SaveChanges: false);
            oDoc = null;
        }


        ////Print to pdf
        //oDoc.PrintOut(ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        //     ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        //     ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        //     ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        //     ref oMissing, ref oMissing);

        //    // Always close Word.exe.
        //    oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
    }
}