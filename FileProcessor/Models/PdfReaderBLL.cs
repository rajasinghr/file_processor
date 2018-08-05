using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Text;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Word;

namespace FileProcessor.Models
{
    public class TextWithFontExtractionStategy : iTextSharp.text.pdf.parser.ITextExtractionStrategy
    {
        private StringBuilder result = new StringBuilder();

        private Vector lastBaseLine;
        private string lastFont;
        private float lastFontSize;

        private enum TextRenderMode
        {
            FillText = 0,
            StrokeText = 1,
            FillThenStrokeText = 2,
            Invisible = 3,
            FillTextAndAddToPathForClipping = 4,
            StrokeTextAndAddToPathForClipping = 5,
            FillThenStrokeTextAndAddToPathForClipping = 6,
            AddTextToPaddForClipping = 7
        }



        public void RenderText(iTextSharp.text.pdf.parser.TextRenderInfo renderInfo)
        {
            string curFont = renderInfo.GetFont().PostscriptFontName;

            if ((renderInfo.GetTextRenderMode() == (int)TextRenderMode.FillThenStrokeText))
            {
                curFont += "-Bold";
            }

            Vector curBaseline = renderInfo.GetBaseline().GetStartPoint();
            Vector topRight = renderInfo.GetAscentLine().GetEndPoint();
            iTextSharp.text.Rectangle rect = new iTextSharp.text.Rectangle(curBaseline[Vector.I1], curBaseline[Vector.I2], topRight[Vector.I1], topRight[Vector.I2]);
            Single curFontSize = rect.Height;

            if ((this.lastBaseLine == null) || (curBaseline[Vector.I2] != lastBaseLine[Vector.I2]) || (curFontSize != lastFontSize) || (curFont != lastFont))
            {
                if ((this.lastBaseLine != null))
                {
                    this.result.AppendLine("\"},");
                }
                if ((this.lastBaseLine != null) && curBaseline[Vector.I2] != lastBaseLine[Vector.I2])
                {
                    this.result.Append("<br />");
                }
                if (this.result.Length == 0)
                {
                    this.result.Append("{\"fontName\":\"" + curFont + "\",\"fontSize\":\"" + curFontSize + "\",\"text\":\"");
                }
                else
                    this.result.Append("{\"fontName\":\"" + curFont + "\",\"fontSize\":\"" + curFontSize + "\",\"text\":\"");
            }

            this.result.Append(renderInfo.GetText());
            this.lastBaseLine = curBaseline;
            this.lastFontSize = curFontSize;
            this.lastFont = curFont;
        }

        public string GetResultantText()
        {
            if (result.Length > 0)
            {
                result.Append("\"},");
            }
            return result.ToString();
        }

        public void BeginTextBlock() { }
        public void EndTextBlock() { }
        public void RenderImage(ImageRenderInfo renderInfo) { }
    }

    public class PdfReaderBLL
    {
        public string createDocument(string path, string docPath)
        {
            Microsoft.Office.Interop.Word.Application winword = null;
            Microsoft.Office.Interop.Word.Document document = null;
            PdfReader reader = null;
            try
            {
                reader = new PdfReader(path);
                List<String> contentsList = new List<string>();
                for (int p = 1; p <= reader.NumberOfPages; p++)
                {
                    TextWithFontExtractionStategy strategy = new TextWithFontExtractionStategy();
                    string resultantString = PdfTextExtractor.GetTextFromPage(reader, p, strategy);
                    string[] dataContents = resultantString.Split(new string[] { "<br />" }, StringSplitOptions.None);
                    foreach (string str in dataContents)
                    {
                        contentsList.Add(str);
                    }
                }


                object missing = System.Reflection.Missing.Value;
                winword = new Microsoft.Office.Interop.Word.Application();
                winword.ShowAnimation = false;
                winword.Visible = false;
                document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                object oStart = 0;
                object oEnd = 0;
                int charCount = 0;
                int oldCharCount = 0;
                foreach (string str in contentsList)
                {
                    string[] lineContent = str.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                    if (lineContent.Length == 1 || (lineContent.Length == 2 && lineContent[1].Length == 0))
                    {
                        TextContent textContentListFull = JsonConvert.DeserializeObject<TextContent>(lineContent[0].ToString().Substring(0, lineContent[0].IndexOf('}') + 1));
                        Paragraph oPara = document.Content.Paragraphs.Add(ref missing);
                        string replacedText = textContentListFull.Text.Replace('\r', ' ').TrimEnd();
                        oPara.Range.Text = replacedText;// +Environment.NewLine;
                        oStart = charCount;
                        oEnd = (int)oStart + replacedText.Length;
                        charCount += replacedText.Length + 1;
                        Range range = document.Range(ref oStart, ref oEnd);
                        range.Font.Size = float.Parse(textContentListFull.FontSize);
                        if (textContentListFull.FontName.IndexOf('+') == 6)
                        {
                            string font = textContentListFull.FontName.Split('+')[1];
                            if (font.Contains('-'))
                            {
                                if (textContentListFull.FontName.Split('+')[1].Split('-')[1].Contains("Bold"))
                                {
                                    range.Font.Bold = 10;
                                    range.Font.Name = textContentListFull.FontName.Split('+')[1].Split('-')[0];
                                }
                                if (textContentListFull.FontName.Split('+')[1].Split('-')[1].Contains("Italic"))
                                {
                                    range.Font.Italic = 10;
                                    range.Font.Name = textContentListFull.FontName.Split('+')[1].Split('-')[0];
                                }
                            }
                            else
                            {
                                range.Font.Name = textContentListFull.FontName.Split('+')[1];
                            }
                        }
                        else
                        {
                            if (textContentListFull.FontName.Substring(textContentListFull.FontName.Length - 2) == "MT")
                            {
                                range.Font.Name = textContentListFull.FontName.Substring(0, textContentListFull.FontName.Length - 2);
                            }
                            else
                            {
                                range.Font.Name = textContentListFull.FontName;
                            }
                        }
                        range.InsertParagraphAfter();
                    }
                    else
                    {
                        List<TextContent> textContents = new List<TextContent>();
                        string concatText = "";
                        for (int i = 0; i < lineContent.Length; i++)
                        {
                            if (lineContent[i].Length > 0)
                            {
                                TextContent textContentListFull = JsonConvert.DeserializeObject<TextContent>(lineContent[i].ToString().Substring(0, lineContent[i].IndexOf('}') + 1));
                                textContents.Add(textContentListFull);
                            }
                        }
                        foreach (TextContent textContent in textContents)
                        {
                            concatText += textContent.Text;
                        }
                        string replacedText = concatText.Replace('\r', ' ').TrimEnd();
                        Paragraph oPara = document.Content.Paragraphs.Add(ref missing);
                        oPara.Range.Text = replacedText;
                        oldCharCount = charCount;
                        charCount += replacedText.Length + 1;
                        foreach (TextContent textContent in textContents)
                        {
                            oStart = oldCharCount;
                            oEnd = (int)oStart + textContent.Text.Length;
                            oldCharCount = (int)oEnd;
                            Range range = document.Range(ref oStart, ref oEnd);
                            range.Font.Size = float.Parse(textContent.FontSize);
                            if (textContent.FontName.IndexOf('+') == 6)
                            {
                                string font = textContent.FontName.Split('+')[1];
                                if (font.Contains('-'))
                                {
                                    if (textContent.FontName.Split('+')[1].Split('-')[1].Contains("Italic"))
                                    {
                                        range.Font.Italic = 10;
                                        range.Font.Name = textContent.FontName.Split('+')[1].Split('-')[0];
                                    }
                                    if (textContent.FontName.Split('+')[1].Split('-')[1].Contains("Bold"))
                                    {
                                        range.Font.Bold = 10;
                                        range.Font.Name = textContent.FontName.Split('+')[1].Split('-')[0];
                                    }
                                }
                                else
                                {
                                    range.Font.Name = textContent.FontName.Split('+')[1];
                                }
                            }
                            else
                            {
                                if (textContent.FontName.Substring(textContent.FontName.Length - 2) == "MT")
                                {
                                    range.Font.Name = textContent.FontName.Substring(0, textContent.FontName.Length - 2);
                                }
                                else
                                {
                                    range.Font.Name = textContent.FontName;
                                }
                            }
                        }
                        oPara.Range.InsertParagraphAfter();
                    }

                }
                //Save the document
                object filename = docPath;
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                reader.Close();
                return docPath;

            }
            catch(Exception ex)
            {
                document = null;
                winword = null;
                reader.Close();
                throw;
            }
        }
    }
}