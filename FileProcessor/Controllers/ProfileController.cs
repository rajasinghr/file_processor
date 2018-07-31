using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace FileProcessor.Controllers
{
    public class ProfileController : Controller
    {
        // GET: Profile
        public ActionResult UserProfile()
        {
            return View();
        }
    }
}