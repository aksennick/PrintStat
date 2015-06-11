using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class CartridgeColorController : BaseController
    {
        //
        // GET: /CartridgeColor/

        public ActionResult Index()
        {
            var cartridgecolor = Repository.CartridgeColors.ToList();
            return View(cartridgecolor);
        }

        [HttpGet]
        public ActionResult CreateCartridgeColor()
        {
            var newcartridgecolorView = new CartridgeColorView();
            return View(newcartridgecolorView);
        }

        [HttpPost]
        public ActionResult CreateCartridgeColor(CartridgeColorView cartridgecolorView)
        {
            var anycartridgecolor = Repository.CartridgeColors.Any(p => string.Compare(p.Name, cartridgecolorView.Name) == 0);
            if (anycartridgecolor)
            {
                ModelState.AddModelError("Name", "Цвет картриджа с таким наименованием уже существует");
            }

            anycartridgecolor = Repository.CartridgeColors.Any(p => string.Compare(p.ShortName, cartridgecolorView.ShortName) == 0);
            if (anycartridgecolor)
            {
                ModelState.AddModelError("Name", "Цвет картриджа с таким коротким наименованием уже существует");
            }

            if (ModelState.IsValid)
            {

                var cartridgecolor = (CartridgeColor)ModelMapper.Map(cartridgecolorView, typeof(CartridgeColorView), typeof(CartridgeColor));
                Repository.CreateCartridgeColor(cartridgecolor);
                return RedirectToAction("Index");
            }

            return View(cartridgecolorView);
        }

        [HttpGet]
        public ActionResult EditCartridgeColor(int? id)
        {
            var cartridgecolor = Repository.CartridgeColors.FirstOrDefault(p => p.ID == id);
            if (cartridgecolor != null)
            {
                var cartridgecolorView = (CartridgeColorView)ModelMapper.Map(cartridgecolor, typeof(CartridgeColor), typeof(CartridgeColorView));
                return View(cartridgecolorView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditCartridgeColor(CartridgeColorView cartridgecolorView)
        {
            var anycartridgecolor = Repository.CartridgeColors.Where(p=>p.ID!=cartridgecolorView.ID).Any(p => string.Compare(p.Name, cartridgecolorView.Name) == 0);
            if (anycartridgecolor)
            {
                ModelState.AddModelError("Name", "Цвет картриджа с таким наименованием уже существует");
            }
            if (ModelState.IsValid)
            {
                var cartridgecolor = Repository.CartridgeColors.FirstOrDefault(p => p.ID == cartridgecolorView.ID);
                ModelMapper.Map(cartridgecolorView, cartridgecolor, typeof(CartridgeColorView), typeof(CartridgeColor));
                Repository.UpdateCartridgeColor(cartridgecolor);

                return RedirectToAction("Index");
            }

            return View(cartridgecolorView);
        }


        [HttpGet]
        public ActionResult DeleteCartridgeColor(int? id)
        {
            var cartridgecolor = Repository.CartridgeColors.FirstOrDefault(p => p.ID == id);
            if (cartridgecolor != null)
            {
                if(!Repository.RemoveCartridgeColor(cartridgecolor))
                {
                    ViewBag.Message = "Невозможно удалить значение, т.к. оно используется";
                    return View("~/Views/Shared/ErrorView.cshtml");
                }
            }
            return RedirectToAction("Index");
        }

    }
}
