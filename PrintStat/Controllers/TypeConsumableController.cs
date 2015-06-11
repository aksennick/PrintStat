using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class TypeConsumableController : BaseController
    {
        //
        // GET: /CartridgeColor/

        public ActionResult Index()
        {
            var _TypeConsumable = Repository.TypeConsumables.ToList();
            return View(_TypeConsumable);
        }

        [HttpGet]
        public ActionResult CreateTypeConsumable()
        {
            var newTypeConsumableView = new TypeConsumableView();
            return View(newTypeConsumableView);
        }

        [HttpPost]
        public ActionResult CreateTypeConsumable(TypeConsumableView _TypeConsumableView)
        {
            var anyTypeConsumable = Repository.TypeConsumables.Any(p => string.Compare(p.Name, _TypeConsumableView.Name) == 0);
            if (anyTypeConsumable)
            {
                ModelState.AddModelError("Name", "Тип компонента с таким наименованием уже существует");
            }

            if (ModelState.IsValid)
            {

                var _TypeConsumable = (TypeConsumable)ModelMapper.Map(_TypeConsumableView, typeof(TypeConsumableView), typeof(TypeConsumable));
                Repository.CreateTypeConsumable(_TypeConsumable);
                return RedirectToAction("Index");
            }

            return View(_TypeConsumableView);
        }

        [HttpGet]
        public ActionResult EditTypeConsumable(int? id)
        {
            var _TypeConsumable = Repository.TypeConsumables.FirstOrDefault(p => p.ID == id);
            if (_TypeConsumable != null)
            {
                var _TypeConsumableView =
                    (TypeConsumableView)ModelMapper.Map(_TypeConsumable, typeof(TypeConsumable), typeof(TypeConsumableView));
                return View(_TypeConsumableView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditTypeConsumable(TypeConsumableView _TypeConsumableView)
        {
            var anyTypeConsumable = Repository.TypeConsumables.Where(p=>p.ID!=_TypeConsumableView.ID).Any(p => string.Compare(p.Name, _TypeConsumableView.Name) == 0);
            if (anyTypeConsumable)
            {
                ModelState.AddModelError("Name", "Тип компонента с таким наименованием уже существует");
            }
            if (ModelState.IsValid)
            {
                var _TypeConsumable = Repository.TypeConsumables.FirstOrDefault(p => p.ID == _TypeConsumableView.ID);
                ModelMapper.Map(_TypeConsumableView, _TypeConsumable, typeof(TypeConsumableView), typeof(TypeConsumable));
                Repository.UpdateTypeConsumable(_TypeConsumable);

                return RedirectToAction("Index");
            }

            return View(_TypeConsumableView);
        }


        [HttpGet]
        public ActionResult DeleteTypeConsumable(int? id)
        {
            var _TypeConsumable = Repository.TypeConsumables.FirstOrDefault(p => p.ID == id);
            if (_TypeConsumable != null)
            {
                if (!Repository.RemoveTypeConsumable(_TypeConsumable))
                {
                    ViewBag.Message = "Невозможно удалить значение, т.к. оно используется";
                    return View("~/Views/Shared/ErrorView.cshtml");
                }
            }
            return RedirectToAction("Index");
        }

    }
}
        
