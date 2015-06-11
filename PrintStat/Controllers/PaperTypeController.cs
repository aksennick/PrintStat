using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class PaperTypeController : BaseController
    {
        //
        // GET: /Paper/

        public ActionResult Index()
        {
            var papertypes = Repository.PaperTypes.ToList();
            return View(papertypes);
        }

        [HttpGet]
        public ActionResult CreatePaperType()
        {
            var newPaperView = new PaperTypeView ();
            return View(newPaperView);
        }


        [HttpPost]
        public ActionResult CreatePaperType(PaperTypeView  papertypeView)
        {
            var anyPaperType = Repository.PaperTypes.Any(p => string.Compare(p.Name, papertypeView.Name)==0);  
            if (anyPaperType)
            {
                ModelState.AddModelError("Name", "Тип бумаги с таким наименованием уже существует");
            }

            if (ModelState.IsValid) 
            {
                
                var paperType = (PaperType)ModelMapper.Map(papertypeView, typeof(PaperTypeView), typeof(PaperType));
                Repository.CreatePapertype(paperType);
                return RedirectToAction("Index");
            }

            return View(papertypeView);
        }


        [HttpGet]
        public ActionResult EditPaperType(int? id)
        {
            var papertype = Repository.PaperTypes.FirstOrDefault(p => p.ID == id);
            if (papertype != null)
            {
                var papertypeView = (PaperTypeView)ModelMapper.Map(papertype, typeof(PaperType), typeof(PaperTypeView));
                return View(papertypeView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditPaperType(PaperTypeView papertypeView)
        {
            var anyPaperType = Repository.PaperTypes.Where(p=>p.ID!=papertypeView.ID).Any(p => string.Compare(p.Name, papertypeView.Name) == 0);
            if (anyPaperType)
            {
                ModelState.AddModelError("Name", "Тип бумаги с таким наименованием уже существует");
            }
            if (ModelState.IsValid)
            {
                var paperType = Repository.PaperTypes.FirstOrDefault(p => p.ID == papertypeView.ID);
                ModelMapper.Map(papertypeView, paperType, typeof(PaperTypeView), typeof(PaperType));
                Repository.UpdatePapertype(paperType);

                return RedirectToAction("Index");
            }

            return View(papertypeView);
        }


        [HttpGet]
        public ActionResult DeletePaperType(int? id)
        {
            var papertype = Repository.PaperTypes.FirstOrDefault(p => p.ID == id);
            if (papertype != null)
            {
                if (!Repository.RemovePapertype(papertype))
                {
                    ViewBag.Message = "Невозможно удалить значение, т.к. оно используется";
                    return View("~/Views/Shared/ErrorView.cshtml");
                }
            }
            return RedirectToAction("Index");
        }

    }
}
