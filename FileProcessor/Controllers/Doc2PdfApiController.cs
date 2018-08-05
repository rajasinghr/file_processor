using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.IO;
using Microsoft.Office.Interop.Word;
using FileProcessor.Models;

namespace FileProcessor.Controllers
{
    public class Doc2PdfApiController : ApiController
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
            for (int cnt = 0; cnt < fileNameArray.Length - 1; cnt++)
            {
                fileNameCorrected += fileNameArray[cnt];
            }
            string fileExtension = "";
            string[] fileExtnArray = fileNameArray[fileNameArray.Length - 1].Split('.');
            fileExtension = fileExtnArray[fileExtnArray.Length - 1];

            httpResponseMessage.Content.Headers.ContentDisposition.FileName = fileNameCorrected + '.' + fileExtension;
            return httpResponseMessage;
        }
        
        [HttpPost()]
        public string Upload()
        {
            string wordFileName = null;
            string wordFileExtension = null;
            string wordFileFullName = null;
            string pdfFileName = null;
            string path = null;
            Application wordApp = null;
            Document wordDoc = null;
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
                if(files.Count == 0)
                {
                    return "Choose a .doc/.docx file";
                }
                for (int count = 0; count <= files.Count - 1; count++)
                {
                    System.Web.HttpPostedFile file = files[count];
                    wordFileName = Path.GetFileNameWithoutExtension(file.FileName)+ "_"+DateTime.Now.ToString("yyyyMMddHHmmssfff");
                    wordFileExtension = Path.GetExtension(file.FileName);
                    wordFileFullName = wordFileName + wordFileExtension;
                    pdfFileName = wordFileName + ".pdf";
                    path = Path.Combine(uploadPath, wordFileFullName);
                    if (file.ContentLength > 0)
                    {
                        file.SaveAs(path);
                        wordApp = new Application();
                        wordDoc = new Document();
                        if (wordFileExtension == ".doc" || wordFileExtension == ".docx")
                        {
                            if (!IsFileLocked(path))
                            {
                                wordDoc = wordApp.Documents.Open(path);
                                wordDoc.ExportAsFixedFormat(Path.Combine(convertPath, pdfFileName), WdExportFormat.wdExportFormatPDF);
                            }
                        }
                        else
                        {
                            throw new FormatException();
                        }
                        wordDoc.Close();
                    }
                }
                return pdfFileName;
            }
            catch(FormatException e)
            {
                ErrorLogging.SendErrorToText(e);
                return "Invalid File format";
            }
            catch(Exception e)
            {
                ErrorLogging.SendErrorToText(e);
                return "Error while uploading file";
            }
        }

        protected bool IsFileLocked(string filePath)
        {
            FileStream stream = null;
            bool locked = false;
            try
            {
                stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                locked = true;
                return locked;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return locked;
        }


    }
}
