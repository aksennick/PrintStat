using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class ComponentController : BaseController
    {
        //
        // GET: /CartridgeColor/

        public ActionResult Index()
        {
            var _Component = Repository.Components.ToList();
            return View(_Component);
        }

        [HttpGet]
        public ActionResult CreateComponent()
        {
            var newComponentView = new ComponentView();
            return View(newComponentView);
        }

        [HttpPost]
        public ActionResult CreateComponent(ComponentView _ComponentView)
        {
            var anyComponent = Repository.Components.Any(p => string.Compare(p.Name, _ComponentView.Name) == 0);
            if (anyComponent)
            {
                ModelState.AddModelError("Name", "Component с таким наименованием уже существует");
            }

            if (ModelState.IsValid)
            {

                var _Component = (Component) ModelMapper.Map(_ComponentView, typeof (ComponentView), typeof (Component));
                Repository.CreateComponent(_Component);
                return RedirectToAction("Index");
            }

            return View(_ComponentView);
        }

        [HttpGet]
        public ActionResult EditComponent(int? id)
        {
            var _Component = Repository.Components.FirstOrDefault(p => p.ID == id);
            if (_Component != null)
            {
                var _ComponentView =
                    (ComponentView) ModelMapper.Map(_Component, typeof (Component), typeof (ComponentView));
                return View(_ComponentView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditComponent(ComponentView _ComponentView)
        {
            if (ModelState.IsValid)
            {
                var _Component = Repository.Components.FirstOrDefault(p => p.ID == _ComponentView.ID);
                ModelMapper.Map(_ComponentView, _Component, typeof (ComponentView), typeof (Component));
                Repository.UpdateComponent(_Component);

                return RedirectToAction("Index");
            }

            return View(_ComponentView);
        }


        [HttpGet]
        public ActionResult DeleteComponent(int? id)
        {
            var _Component = Repository.Components.FirstOrDefault(p => p.ID == id);
            if (_Component != null)
            {
                Repository.RemoveComponent(_Component);
            }
            return RedirectToAction("Index");
        }

    }
}
        
