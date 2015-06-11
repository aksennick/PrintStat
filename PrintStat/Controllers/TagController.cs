using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{

    public class TagController : BaseController
    {
        //
        // GET: /Tag/
        private void InitViewBag()
        {
            ViewBag.TagTypes = Repository.TagTypes;
        }
        public ActionResult Index()
        {
            var _Tag = Repository.Tags.ToList();
            return View(_Tag);
        }

        [HttpGet]
        public ActionResult CreateTag()
        {
            InitViewBag();
            var newTagView = new TagView();
            return View(newTagView);
        }

        [HttpPost]
        public ActionResult CreateTag(TagView _TagView)
        {
            var anyTag = Repository.Tags.Any(p => string.Compare(p.Name, _TagView.Name) == 0);
            if (anyTag)
            {
                ModelState.AddModelError("Name", "Тег с таким наименованием уже существует");
            }
            anyTag = Repository.Tags.Any(p => string.Compare(p.Tag1, _TagView.Tag1) == 0);
            if (anyTag)
            {
                ModelState.AddModelError("Tag1", "Такой тег уже существует");
            }

            if (ModelState.IsValid)
            {

                var _Tag = (Tag)ModelMapper.Map(_TagView, typeof(TagView), typeof(Tag));
                Repository.CreateTag(_Tag);
                return RedirectToAction("Index");
            }

            return View(_TagView);
        }

        [HttpGet]
        public ActionResult EditTag(int? id)
        {
            InitViewBag();
            var _Tag = Repository.Tags.FirstOrDefault(p => p.ID == id);
            if (_Tag != null)
            {
                var _TagView = (TagView)ModelMapper.Map(_Tag, typeof(Tag), typeof(TagView));
                return View(_TagView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditTag(TagView _TagView)
        {
            if (ModelState.IsValid)
            {
                var _Tag = Repository.Tags.FirstOrDefault(p => p.ID == _TagView.ID);
                ModelMapper.Map(_TagView, _Tag, typeof(TagView), typeof(Tag));
                Repository.UpdateTag(_Tag);

                return RedirectToAction("Index");
            }

            return View(_TagView);
        }


        [HttpGet]
        public ActionResult DeleteTag(int? id)
        {
            var _Tag = Repository.Tags.FirstOrDefault(p => p.ID == id);
            if (_Tag != null)
            {
                if (!Repository.RemoveTag(_Tag))
                {
                    ViewBag.Message = "Невозможно удалить значение, т.к. оно используется";
                    return View("~/Views/Shared/ErrorView.cshtml");
                }
            }
            return RedirectToAction("Index");
        }

    }
        
}
