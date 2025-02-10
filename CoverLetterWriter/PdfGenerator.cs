using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using SelectPdf;
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
            //string baseDirectory = 
            //string filePath = 
            //string startupPath2 = Directory.GetCurrentDirectory();
            //string startupPath = Path.Combine(startupPath2, "CoverLetter.docx");

            //filePath = Directory. + filePath;
            // Create a new instance of Word application
            Application wordApp = new Application();
            string newcompname = string.Empty;
            try
            {
                wordApp.Visible = false;
                string htmlContent = @"<!DOCTYPE html><html lang='de'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Bewerbung</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 40px;
            line-height: 1.6;
        }
        .header {
        display: flex;
        justify-content: space-between;
        }
        .header, .footer {
            margin-bottom: 20px;
        }
        .content {
            margin-top: 20px;
        }
        .signature {
            margin-top: 40px;
        }
    </style>
</head>
<body>
    <div class='header'>
    <p ><br> <br> <br> {CompanyName}<br>
        {CompanyStreetAddress}<br>
        {CompanyCityAndPostcode}</p>
        <p>{FullName}<br>
        {StreetAddress}<br>
        {CityAndPostcode}<br>
        {PhoneNumber}<br>
        {Email} <br> <br> <br> <br>
        {Date}</p>

        

    </div>

    <div class='content'>
        <p><strong>Bewerbung als {PositionFullName}</strong></p>

        <p>Sehr {Name},</p>

        <p>mit großem Interesse habe ich auf Ihrer Website die Ausschreibung für die Position {PositionName} gelesen. Mein Name ist {FullName}, und ich bin begeistert von der Möglichkeit, zum kontinuierlichen Erfolg und Wachstum Ihres Unternehmens beizutragen.</p>

        <p>{MainText}</p>
        <p> Vielen Dank, dass Sie meine Bewerbung in Betracht ziehen. Ich freue mich auf die Gelegenheit, mit Ihnen in Kontakt zu treten.</ p >
    </ div >

    <div class='signature'>
        <p>Mit freundlichen Grüßen,</p>
        <p>{FullName}</p>
    </div>
</body>
</html>";
                var converter = new HtmlToPdf();
                foreach (KeyValue kv in newText)
                {
                    if (kv.Key == "{CompanyName}")
                    {
                        newcompname = kv.Value;
                    }
                    htmlContent =  htmlContent.Replace(kv.Key, kv.Value);

                }
                PdfDocument doc = converter.ConvertHtmlString(htmlContent);


                

                // Find and replace text

                //Find findObj = wordApp.Selection.Find;
               




                doc.Save(filePath);
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