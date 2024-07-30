using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using Application = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word.Document;
//using Document = iTextSharp.text.Document;
//using Paragraph = iTextSharp.text.Paragraph;

namespace CoverLetterWriter
{
    internal class PdfGenerator
    {
        public void EditWordDocument(string filePath, List<KeyValue> newText)
        {
            string startupPath2 = Environment.CurrentDirectory;
            string startupPath = Directory.GetParent(startupPath2).Parent.Parent.FullName + "\\";

            //filePath = Directory. + filePath;
            // Create a new instance of Word application
            Application wordApp = new Application();
            string newcompname = string.Empty;
            try
            {
                wordApp.Visible = false;
                // Open the Word document
                Document doc = wordApp.Documents.Open(startupPath + filePath);

                // Find and replace text
                Find findObj = wordApp.Selection.Find;
                foreach (KeyValue kv in newText)
                {
                    if(kv.Key == "{CompanyName}")
                    {
                        newcompname = kv.Value;
                    }
                    findObj.Text = kv.Key;
                    findObj.Replacement.Text = kv.Value;
                    findObj.Execute(Replace: WdReplace.wdReplaceAll);
                }




                doc.SaveAs2(startupPath + newcompname + "-cover letter.docx");
                doc.Close();
                doc = wordApp.Documents.Open(startupPath + newcompname + "-cover letter.docx");
                // Save the document
                doc.SaveAs2(startupPath + newcompname + "-cover letter.pdf", WdSaveFormat.wdFormatPDF);
             
                // Close the original document
                doc.Close();

            }
            catch (Exception ex)
            {
                // Handle any errors that occur during Word document editing
                //Console.Error.WriteLine("Error editing Word document: " + ex.Message);
                string messageBoxText = "Something Went Wrong";
                string caption = ex.Message;
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBoxImage icon = MessageBoxImage.Warning;
                MessageBoxResult result;

                result = MessageBox.Show(messageBoxText, caption, button, icon, MessageBoxResult.Yes);
            }
            finally
            {
                // Close Word application
                wordApp.Quit();
                string messageBoxText = "Done";
                string caption = "Your pdf is ready";
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBoxImage icon = MessageBoxImage.Information;
                MessageBoxResult result;

                result = MessageBox.Show(messageBoxText, caption, button, icon, MessageBoxResult.Yes);
            }

        }
    }
}