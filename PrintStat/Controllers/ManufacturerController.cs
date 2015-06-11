using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{

    public class ManufacturerController : BaseController
    {
        //
        // GET: /CartridgeColor/

        public ActionResult Index()
        {
            var _Manufacturer = Repository.Manufacturers.ToList();
            return View(_Manufacturer);
        }

        [HttpGet]
        public ActionResult CreateManufacturer()
        {
            var newManufacturerView = new ManufacturerView();
            return View(newManufacturerView);
        }

        [HttpPost]
        public ActionResult CreateManufacturer(ManufacturerView _ManufacturerView)
        {
            var anyManufacturer = Repository.Manufacturers.Any(p => string.Compare(p.Name, _ManufacturerView.Name) == 0);
            if (anyManufacturer)
            {
                ModelState.AddModelError("Name", "Производитель с таким наименованием уже существует");
            }

            if (ModelState.IsValid)
            {

                var _Manufacturer =
                    (Manufacturer) ModelMapper.Map(_ManufacturerView, typeof (ManufacturerView), typeof (Manufacturer));
                Repository.CreateManufacturer(_Manufacturer);
                return RedirectToAction("Index");
            }

            return View(_ManufacturerView);
        }

        [HttpGet]
        public ActionResult EditManufacturer(int? id)
        {
            var _Manufacturer = Repository.Manufacturers.FirstOrDefault(p => p.ID == id);
            if (_Manufacturer != null)
            {
                var _ManufacturerView =
                    (ManufacturerView) ModelMapper.Map(_Manufacturer, typeof (Manufacturer), typeof (ManufacturerView));
                return View(_ManufacturerView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditManufacturer(ManufacturerView _ManufacturerView)
        {
            var anyManufacturer = Repository.Manufacturers.Where(p=>p.ID!=_ManufacturerView.ID).Any(p => string.Compare(p.Name, _ManufacturerView.Name) == 0);
            if (anyManufacturer)
            {
                ModelState.AddModelError("Name", "Производитель с таким наименованием уже существует");
            }
            if (ModelState.IsValid)
            {
                var _Manufacturer = Repository.Manufacturers.FirstOrDefault(p => p.ID == _ManufacturerView.ID);
                ModelMapper.Map(_ManufacturerView, _Manufacturer, typeof (ManufacturerView), typeof (Manufacturer));
                Repository.UpdateManufacturer(_Manufacturer);

                return RedirectToAction("Index");
            }

            return View(_ManufacturerView);
        }


        [HttpGet]
        public ActionResult DeleteManufacturer(int? id)
        {
            var _Manufacturer = Repository.Manufacturers.FirstOrDefault(p => p.ID == id);
            if (_Manufacturer != null)
            {
                if (!Repository.RemoveManufacturer(_Manufacturer))
                {
                    ViewBag.Message = "Невозможно удалить значение, т.к. оно используется";
                    return View("~/Views/Shared/ErrorView.cshtml");
                }
            }
            return RedirectToAction("Index");
        }


        public JsonResult AutocompleteManufacturer(string term)
        {

            return Json(Repository.SearchManufacturer(term), JsonRequestBehavior.AllowGet);

        }

        [AcceptVerbs(HttpVerbs.Post)]
        public int? AddManufacturer(string name)
        {
            if (name == null) return null;
            int? id = Repository.CheckManufacturer(name);
            if (id == 0)
            {
                Manufacturer insence = new Manufacturer()
                {
                    Name = name
                };
                Repository.CreateManufacturer(insence);
                id = insence.ID;
            }
            return id;
        }





    }
}

