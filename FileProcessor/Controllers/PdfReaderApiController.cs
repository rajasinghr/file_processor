using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Text;
using System.IO;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;

namespace FileProcessor.Controllers
{
    public class PdfReaderApiController : ApiController
    {

        [HttpGet]
        public HttpResponseMessage Download(string fileName)
        {
            HttpResponseMessage httpResponseMessage = null;
            string localFilePath = System.Web.Hosting.HostingEnvironment.MapPath("~/ConvertedFiles/");
            httpResponseMessage = Request.CreateResponse(HttpStatusCode.OK);
            httpResponseMessage.Content = new StreamContent(new FileStream(localFilePath + fileName, FileMode.Open, FileAccess.Read));
            httpResponseMessage.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
            string[] fileNameArray = fileName.Split('_');
            string fileNameCorrected = "";
            for(int cnt=0;cnt<fileNameArray.Length-1;cnt++)
            {
                fileNameCorrected += fileNameArray[cnt];
            }
            string fileExtension = "";
            string[] fileExtnArray = fileNameArray[fileNameArray.Length-1].Split('.');
            fileExtension = fileExtnArray[fileExtnArray.Length-1];

            httpResponseMessage.Content.Headers.ContentDisposition.FileName = fileNameCorrected+'.'+ fileExtension;
            return httpResponseMessage;
        }


        [HttpPost]
        public String Pdf2Doc()
        {
            string pdfFileName = null;
            string pdfFileExtension = null;
            string pdfFullFileName = null;
            string docFileName = null;
            string path = null;
            string docPath = null;
            string uploadPath = "";
            string convertPath = "";
            uploadPath = System.Web.Hosting.HostingEnvironment.MapPath("~/UploadedFiles/");
            convertPath = System.Web.Hosting.HostingEnvironment.MapPath("~/ConvertedFiles/");
            StringBuilder contents = null;
            if (!Directory.Exists(uploadPath))
            {
                Directory.CreateDirectory(uploadPath);
            }
            if (!Directory.Exists(convertPath))
            {
                Directory.CreateDirectory(convertPath);
            }
            try
            {
                System.Web.HttpFileCollection files = System.Web.HttpContext.Current.Request.Files;
                var option = System.Web.HttpContext.Current.Request.Params["radioOption"];
                if (files.Count == 0)
                {
                    return "Choose a .pdf file";
                }
                for (int count = 0; count <= files.Count - 1; count++)
                {
                    System.Web.HttpPostedFile file = files[count];

                    if (file.ContentLength > 0)
                    {

                        pdfFileName = System.IO.Path.GetFileNameWithoutExtension(file.FileName)+ "_"+DateTime.Now.ToString("yyyyMMddHHmmssfff");
                        pdfFileExtension = System.IO.Path.GetExtension(file.FileName);
                        if (pdfFileExtension != ".pdf")
                        {
                            throw new FormatException();
                        }
                        pdfFullFileName = pdfFileName + pdfFileExtension;
                        if (option == ".doc")
                        {
                            docFileName = pdfFileName + ".doc";
                        }
                        else if (option == ".docx")
                        {
                            docFileName = pdfFileName + ".docx";
                        }
                        path = System.IO.Path.Combine(uploadPath, pdfFullFileName);
                        docPath = System.IO.Path.Combine(convertPath, docFileName);
                        file.SaveAs(path);
                        var pdf = new Aspose.Pdf.Document(path);
                        if (option == ".doc")
                        {
                            pdf.Save(docPath, Aspose.Pdf.SaveFormat.Doc);
                        }
                        else if (option == ".docx")
                        {
                            pdf.Save(docPath, Aspose.Pdf.SaveFormat.DocX);
                        }
                    }
                }
                return docFileName;
            }
            catch (FormatException)
            {
                return "Invalid File format";
            }
            catch (Exception e)
            {
                return "Error Occured";
            }
        }

        [HttpPost]
        public String Pdf2Text()
        {
            string pdfFileName = null;
            string pdfFileExtension = null;
            string pdfFullFileName = null;
            string textFileName = null;
            string path = null;
            string textPath = null;
            string uploadPath = "";
            string convertPath = "";
            uploadPath = System.Web.Hosting.HostingEnvironment.MapPath("~/UploadedFiles/");
            convertPath = System.Web.Hosting.HostingEnvironment.MapPath("~/ConvertedFiles/");
            if (!Directory.Exists(uploadPath))
            {
                Directory.CreateDirectory(uploadPath);
            }
            if (!Directory.Exists(convertPath))
            {
                Directory.CreateDirectory(convertPath);
            }
            try
            {
                System.Web.HttpFileCollection files = System.Web.HttpContext.Current.Request.Files;
                var option = System.Web.HttpContext.Current.Request.Params["radioOption"];
                if (files.Count == 0)
                {
                    return "Choose a .pdf file";
                }
                for (int count = 0; count <= files.Count - 1; count++)
                {
                    System.Web.HttpPostedFile file = files[count];

                    if (file.ContentLength > 0)
                    {
                        pdfFileName = System.IO.Path.GetFileNameWithoutExtension(file.FileName)+"_"+ DateTime.Now.ToString("yyyyMMddHHmmssfff");
                        pdfFileExtension = System.IO.Path.GetExtension(file.FileName);
                        if(pdfFileExtension!=".pdf")
                        {
                            throw new FormatException();
                        }
                        pdfFullFileName = pdfFileName + pdfFileExtension;
                        textFileName = pdfFileName + ".txt";
                        path = System.IO.Path.Combine(uploadPath, pdfFullFileName);
                        textPath = System.IO.Path.Combine(convertPath, textFileName);
                        file.SaveAs(path);

                        ITextExtractionStrategy its = new LocationTextExtractionStrategy();
                        PdfReader reader = new PdfReader(path);
                        using (StreamWriter writer = new StreamWriter(textPath))
                        {
                            for (int i = 1; i <= reader.NumberOfPages; i++)
                            {
                                string thePage = PdfTextExtractor.GetTextFromPage(reader, i, its);
                                string[] theLines = thePage.Split('\n');
                                foreach (var theLine in theLines)
                                {
                                    writer.WriteLine(theLine);
                                }
                            }
                        }
                    }
                }
                return textFileName;
            }
            catch (FormatException)
            {
                return "Invalid File format";
            }
            catch (Exception)
            {
                return "Error Occured";
            }
        }
    }
}