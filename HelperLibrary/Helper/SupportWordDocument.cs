using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using HelperLibrary.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary
{
    public class SupportWordDocument
    {
        public static int ReplaceParameters(string documentFileName, Dictionary<string, string> parameters, params SupportWordTabelModel[] items)
        {
            int count = 0;
            string parameterName = null;
            List<Text> parameterTexts = new List<Text>();
            int icount = 1;
            using (DocumentFormat.OpenXml.Packaging.WordprocessingDocument document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(documentFileName, true))
            {
                var body = document.MainDocumentPart.Document.Body;

                foreach (SupportWordTabelModel item in items)
                {
                    if (item.Data != null)
                    {
                        foreach (Table t in body.Descendants<Table>().Where(tbl => tbl.InnerText.Contains(item.TableKeyword)))
                        {
                            foreach (List<string> tableLine in item.Data)
                            {
                                TableRow row = new TableRow();
                                TableCell no = new TableCell(new Paragraph(new Run(new Text(icount.ToString()))));
                                row.Append(no);
                                foreach (string tabelElement in tableLine)
                                {
                                    TableCell cell = new TableCell(new Paragraph(new Run(new Text(tabelElement))));
                                    row.Append(cell);
                                }

                                t.Append(row);
                                icount++;
                            }
                        }
                    }

                }

                // Process all paragraphs
                foreach (var para in body.Elements<Paragraph>())
                {
                    count += ParseParagraph(para, parameters, ref parameterName, ref parameterTexts);
                }

                // Process all tables
                foreach (var table in body.Elements<Table>())
                {
                    foreach (var row in table.Elements<TableRow>())
                    {
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            // Process all paragraphs
                            foreach (var para in cell.Elements<Paragraph>())
                            {
                                count += ParseParagraph(para, parameters, ref parameterName, ref parameterTexts);
                            }
                        }
                    }
                }
            }

            return count;
        }

        private static int ParseParagraph(Paragraph paragraph, Dictionary<string, string> parameters, ref string parameterName, ref List<Text> parameterTexts)
        {
            int count = 0;

            foreach (var run in paragraph.Elements<Run>())
            {
                foreach (var text in run.Elements<Text>())
                {
                    switch (text.Text)
                    {
                        case "←":
                            // Parameter started
                            parameterName = string.Empty;
                            parameterTexts.Clear();
                            parameterTexts.Add(text);
                            break;
                        case "→":
                            // Parameter ended
                            if (parameterName != null)
                            {
                                if (parameters.ContainsKey(parameterName))
                                {
                                    // Replace parameter name with actual value
                                    count++;
                                    string replacement = parameters[parameterName];

                                    foreach (Text text2 in parameterTexts)
                                    {
                                        text2.Text = string.Empty;
                                    }

                                    parameterTexts[0].Text = replacement;

                                    text.Text = string.Empty;
                                }
                            }

                            parameterName = null;
                            parameterTexts.Clear();
                            break;
                        default:
                            if (parameterName != null)
                            {
                                // Parameter name
                                parameterName += text.Text;
                                parameterTexts.Add(text);
                            }
                            else
                            {
                                // Plain text
                            }

                            break;
                    }
                }
            }

            return count;
        }
    }
}
