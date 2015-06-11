using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{

    public class TagTypeController : BaseController
    {
        //
        // GET: /CartridgeColor/

        public ActionResult Index()
        {
            var _TagType = Repository.TagTypes.ToList();
            return View(_TagType);
        }

        [HttpGet]
        public ActionResult CreateTagType()
        {
            var newTagTypeView = new TagTypeView();
            return View(newTagTypeView);
        }

        [HttpPost]
        public ActionResult CreateTagType(TagTypeView _TagTypeView)
        {
            var anyTagType = Repository.TagTypes.Any(p => string.Compare(p.Type, _TagTypeView.Type) == 0);
            if (anyTagType)
            {
                ModelState.AddModelError("Type", "Такой тип уже существует");
            }

            if (ModelState.IsValid)
            {

                var _TagType = (TagType)ModelMapper.Map(_TagTypeView, typeof(TagTypeView), typeof(TagType));
                Repository.CreateTagType(_TagType);
                return RedirectToAction("Index");
            }

            return View(_TagTypeView);
        }

        [HttpGet]
        public ActionResult EditTagType(int? id)
        {
            var _TagType = Repository.TagTypes.FirstOrDefault(p => p.ID == id);
            if (_TagType != null)
            {
                var _TagTypeView = (TagTypeView)ModelMapper.Map(_TagType, typeof(TagType), typeof(TagTypeView));
                return View(_TagTypeView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditTagType(TagTypeView _TagTypeView)
        {
            var anyTagType = Repository.TagTypes.Where(p=>p.ID!=_TagTypeView.ID).Any(p => string.Compare(p.Type, _TagTypeView.Type) == 0);
            if (anyTagType)
            {
                ModelState.AddModelError("Type", "Такой тип уже существует");
            }
            if (ModelState.IsValid)
            {
                var _TagType = Repository.TagTypes.FirstOrDefault(p => p.ID == _TagTypeView.ID);
                ModelMapper.Map(_TagTypeView, _TagType, typeof(TagTypeView), typeof(TagType));
                Repository.UpdateTagType(_TagType);

                return RedirectToAction("Index");
            }

            return View(_TagTypeView);
        }


        [HttpGet]
        public ActionResult DeleteTagType(int? id)
        {
            var _TagType = Repository.TagTypes.FirstOrDefault(p => p.ID == id);
            if (_TagType != null)
            {
                if (!Repository.RemoveTagType(_TagType))
                {
                    ViewBag.Message = "Невозможно удалить значение, т.к. оно используется";
                    return View("~/Views/Shared/ErrorView.cshtml");
                }
            }
            return RedirectToAction("Index");
        }

    }
        
}
