using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class SizePaperController : BaseController
    {
        //
        // GET: /SizePaper/

        public ActionResult Index()
        {
            var papersize = Repository.SizePapers.ToList();
            return View(papersize);
            
        }

        [HttpGet]
        public ActionResult CreateSizePaper()
        {
            var newSizePaperView = new SizePaperView();
            return View(newSizePaperView);
        }


        [HttpPost]
        public ActionResult CreateSizePaper(SizePaperView sizepaperView)
        {
            var anySizePaper = Repository.SizePapers.Any(p => string.Compare(p.Name, sizepaperView.Name) == 0);
            if (anySizePaper)
            {
                ModelState.AddModelError("Name", "Типоразмер бумаги с таким наименованием уже существует");
            }

            if (ModelState.IsValid)
            {
                var sizePaper = (SizePaper)ModelMapper.Map(sizepaperView, typeof(SizePaperView), typeof(SizePaper));
                Repository.CreateSizePaper(sizePaper);
                return RedirectToAction("Index");
            }

            return View(sizepaperView);
        }


        [HttpGet]
        public ActionResult EditSizePaper(int? id)
        {
            var sizepaper = Repository.SizePapers.FirstOrDefault(p => p.ID == id);
            if (sizepaper != null)
            {
                var papersizeView = (SizePaperView)ModelMapper.Map(sizepaper, typeof(SizePaper), typeof(SizePaperView));
                return View(papersizeView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditSizePaper(SizePaperView papersizeView)
        {
            var anySizePaper = Repository.SizePapers.Where(p => p.ID != papersizeView.ID).Any(p => string.Compare(p.Name, papersizeView.Name) == 0);
            if (anySizePaper)
            {
                ModelState.AddModelError("Name", "Типоразмер бумаги с таким наименованием уже существует");
            }
            if (ModelState.IsValid)
            {
                var papersize = Repository.SizePapers.FirstOrDefault(p => p.ID == papersizeView.ID);
                ModelMapper.Map(papersizeView, papersize, typeof(SizePaperView), typeof(SizePaper));
                Repository.UpdateSizePaper(papersize);

                return RedirectToAction("Index");
            }

            return View(papersizeView);
        }


        [HttpGet]
        public ActionResult DeleteSizePaper(int? id)
        {
            var papersize = Repository.SizePapers.FirstOrDefault(p => p.ID == id);
            if (papersize != null)
            {
                if (!Repository.RemoveSizePaper(papersize))
                {
                    ViewBag.Message = "Невозможно удалить значение, т.к. оно используется";
                    return View("~/Views/Shared/ErrorView.cshtml");
                }
            }
            return RedirectToAction("Index");
        }
    }
}
