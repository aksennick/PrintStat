using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class PrintKindController : BaseController
    {
        //
        // GET: /PrintKind/

        public ActionResult Index()
        {
            var printkinds = Repository.PrintKinds.ToList();
            return View(printkinds);
        }

        [HttpGet]
        public ActionResult CreatePrintKind()
        {
            var newPrintKindView = new PrintKindView();
            return View(newPrintKindView);
        }

        [HttpPost]
        public ActionResult CreatePrintKind(PrintKindView printkindView)
        {
            var anyPrintKind = Repository.PrintKinds.Any(p => string.Compare(p.Name, printkindView.Name) == 0);
            if (anyPrintKind)
            {
                ModelState.AddModelError("Name", "Вид печати с таким наименованием уже существует");
            }

            if (ModelState.IsValid)
            {

                var printKind = (PrintKind)ModelMapper.Map(printkindView, typeof(PrintKindView), typeof(PrintKind));
                Repository.CreatePrintKind(printKind);
                return RedirectToAction("Index");
            }

            return View(printkindView);
        }

        [HttpGet]
        public ActionResult EditPrintKind(int? id)
        {
            var printKind = Repository.PrintKinds.FirstOrDefault(p => p.ID == id);
            if (printKind != null)
            {
                var printKindView = (PrintKindView)ModelMapper.Map(printKind, typeof(PrintKind), typeof(PrintKindView));
                return View(printKindView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditPrintKind(PrintKindView printkindView)
        {
            var anyPrintKind = Repository.PrintKinds.Where(p=>p.ID!=printkindView.ID).Any(p => string.Compare(p.Name, printkindView.Name) == 0);
            if (anyPrintKind)
            {
                ModelState.AddModelError("Name", "Вид печати с таким наименованием уже существует");
            }
            if (ModelState.IsValid)
            {
                var printKind= Repository.PrintKinds.FirstOrDefault(p => p.ID == printkindView.ID);
                ModelMapper.Map(printkindView, printKind, typeof(PrintKindView), typeof(PrintKind));
                Repository.UpdatePrintKind(printKind);

                return RedirectToAction("Index");
            }

            return View(printkindView);
        }


        [HttpGet]
        public ActionResult DeletePrintKind(int? id)
        {
            var printkind = Repository.PrintKinds.FirstOrDefault(p => p.ID == id);
            if (printkind != null)
            {
                if (!Repository.RemovePrintKind(printkind))
                {
                    ViewBag.Message = "Невозможно удалить значение, т.к. оно используется";
                    return View("~/Views/Shared/ErrorView.cshtml");
                }
            }
            return RedirectToAction("Index");
        }
    }
}
