using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Assignment_Test.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        //resposible for handling errors
        public ActionResult Error()
        {
            ViewBag.ErrorMessage = "An unexpected error occurred. Please try again later.";
            return View();
        }

        //responsible for handling 404 errors
        public ActionResult NotFound()
        {
           
            return View();
        }
    }
}