using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;
using PagedList;
using PagedList.Mvc;

namespace PrintStat.Controllers
{
    public class JobController : BaseController
    {



        public ActionResult Index()
        {
            var Jobs = Repository.Jobs.ToList();
            return View(Jobs);
        }

        [HttpGet]
        public ActionResult CreateJob()
        {
            var newJobView = new JobView ();
            return View(newJobView);
        }


        [HttpPost]
        public ActionResult CreateJob(JobView  JobView)
        {
            var anyJob = Repository.Jobs.Any(p => string.Compare(p.Name, JobView.Name)==0);  
            if (anyJob)
            {
                ModelState.AddModelError("Name", "Тип бумаги с таким наименованием уже существует");
            }

            if (ModelState.IsValid) 
            {
                var Job = (Job)ModelMapper.Map(JobView, typeof(JobView), typeof(Job));
                Repository.CreateJob(Job);
                return RedirectToAction("Index");
            }

            return View(JobView);
        }


        [HttpGet]
        public ActionResult EditJob(int? id)
        {
            var Job = Repository.Jobs.FirstOrDefault(p => p.ID == id);
            if (Job != null)
            {
                var JobView = (JobView)ModelMapper.Map(Job, typeof(Job), typeof(JobView));
                return View(JobView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditJob(JobView JobView)
        {
            if (ModelState.IsValid)
            {
                var Job = Repository.Jobs.FirstOrDefault(p => p.ID == JobView.ID);
                ModelMapper.Map(JobView, Job, typeof(JobView), typeof(Job));
                Repository.UpdateJob(Job);

                return RedirectToAction("Index");
            }

            return View(JobView);
        }


        [HttpGet]
        public ActionResult DeleteJob(int? id)
        {
            var Job = Repository.Jobs.FirstOrDefault(p => p.ID == id);
            if (Job != null)
            {
                Repository.RemoveJob(Job);
            }
            return RedirectToAction("Index");
        }


        public ActionResult partialJobPaginator(int? page, int? countJob)
        {
            var Jobs = Repository.Jobs.OrderByDescending(s => s.EndTime); ; //returns IQueryable<Product> representing an unknown number of products. a thousand maybe?

            var pageNumber = page ?? 1; // if no page was specified in the querystring, default to the first page (1)
            var count = countJob ?? 30;
            return PartialView("partialJobPaginator", Jobs.ToPagedList(pageNumber, count));
        }



    }
}
