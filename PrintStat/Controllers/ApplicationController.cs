using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class ApplicationController : BaseController
    {

        //
        // GET: /Application/

        public ActionResult Index()
        {
            var applications = Repository.Applications.ToList();
            return View(applications);
        }

        [HttpGet]
        public ActionResult CreateApplication()
        {
            var newapplicationsView = new ApplicationView();
            return View(newapplicationsView);
        }

        [HttpPost]
        public ActionResult CreateApplication(ApplicationView applicationView)
        {

            var anyApplication = Repository.Applications.Any(p => string.Compare(p.Name, applicationView.Name) == 0);
            if (anyApplication)
            {
                ModelState.AddModelError("Name", "Приложение с таким наименованием уже существует");
            }

            if (ModelState.IsValid)
            {

                var application= (Application)ModelMapper.Map(applicationView, typeof(ApplicationView), typeof(Application));
                Repository.CreateApplication(application);
                return RedirectToAction("Index");
            }

            return View(applicationView);
        }

        [HttpGet]
        public ActionResult EditApplication(int? id)
        {
            var application = Repository.Applications.FirstOrDefault(p => p.ID == id);
            if (application != null)
            {
                var applicationView = (ApplicationView)ModelMapper.Map(application, typeof(Application), typeof(ApplicationView));
                return View(applicationView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditApplication(ApplicationView applicationView)
        {
            var anyApplication = Repository.Applications.Where(p=>p.ID!=applicationView.ID).Any(p => string.Compare(p.Name, applicationView.Name) == 0);
            if (anyApplication)
            {
                ModelState.AddModelError("Name", "Приложение с таким наименованием уже существует");
            }
            if (ModelState.IsValid)
            {
                var application = Repository.Applications.FirstOrDefault(p => p.ID == applicationView.ID);
                ModelMapper.Map(applicationView, application, typeof(ApplicationView), typeof(Application));
                Repository.UpdateApplication(application);

                return RedirectToAction("Index");
            }

            return View(applicationView);
        }


        [HttpGet]
        public ActionResult DeleteApplication(int? id)
        {
            var application = Repository.Applications.FirstOrDefault(p => p.ID == id);
            if (application != null)
            {
                if (!Repository.RemoveApplication(application))
                {
                    ViewBag.Message = "Невозможно удалить значение, т.к. оно используется";
                    return View("~/Views/Shared/ErrorView.cshtml");
                }
            }
            return RedirectToAction("Index");
        }
    }
}


