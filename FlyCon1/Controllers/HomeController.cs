using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MvcProject.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Message = "Modify this template to jump-start your ASP.NET MVC application.";

            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your app description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult Slider()
        {
            return View();
        }
        public ActionResult Graph()
        {
            return View();
        }
        public ActionResult InputScreen()
        {
            return View();
        }
        public ActionResult MainPage()
        {
            return View();
        }
        public ActionResult MainPageIcons()
        {
            return View();
        }
        public ActionResult MultiAccordion()
        {
            return View();
        }
       
    }
}
