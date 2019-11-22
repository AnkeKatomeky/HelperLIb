using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary
{
    public static class PDFHelper
    {
        public static List<List<string>> PDFText(string path, bool isSpace = false)
        {
            List<List<string>> textLines = new List<List<string>>();
            string[] textPage;

            using (PdfReader reader = new PdfReader(path))
            {
                string text = string.Empty;
                for (int page = 1; page <= reader.NumberOfPages; page++)
                {
                    if (isSpace)
                    {
                        text += PdfTextExtractor.GetTextFromPage(reader, page, new SimpleTextExtractionStrategy()).Replace(" ", "");
                    }
                    else
                    {
                        text += PdfTextExtractor.GetTextFromPage(reader, page, new SimpleTextExtractionStrategy());
                    }
                    text += "#############################################";
                }
                reader.Close();
                textPage = text.Split(new string[] { "#############################################" }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string item in textPage)
                {
                    textLines.Add(new List<string>());
                    textLines.Last().AddRange(item.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries));
                }
            }
            return textLines;
        }
    }
}
