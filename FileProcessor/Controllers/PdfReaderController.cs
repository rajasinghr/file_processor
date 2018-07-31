using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace FileProcessor.Controllers
{
    public class PdfReaderController : Controller
    {
        public ActionResult Pdf2Doc()
        {
            return View();
        }

        public ActionResult Pdf2Text()
        {
            return View();
        }
        
    }
}