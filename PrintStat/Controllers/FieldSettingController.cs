using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class FieldSettingController : BaseController
    {
        //
        // GET: /FieldSetting/

        public ActionResult Index()
        {
            var fieldSetts = Repository.FieldSettings.ToList();
            return View(fieldSetts);
        }


        [HttpGet]
        public ActionResult CreateFieldSetting()
        {
            var newfieldView = new FieldSettingView();
            return View(newfieldView);
        }

        [HttpPost]
        public ActionResult CreateFieldSetting(FieldSettingView fieldView)
        {
            var anyfield = Repository.FieldSettings.Any(p => string.Compare(p.Name, fieldView.Name) == 0);
            if (anyfield)
            {
                ModelState.AddModelError("ParameterName", "Такое поле настройки уже существует");
            }

            if (ModelState.IsValid)
            {

                var fieldsetting = (Settings)ModelMapper.Map(fieldView, typeof(FieldSettingView), typeof(Settings));
                Repository.CreateFieldSetting(fieldsetting);
                return RedirectToAction("Index");
            }

            return View(fieldView);
        }

        [HttpGet]
        public ActionResult EditFieldSetting(int? id)
        {
            var fieldsetting = Repository.FieldSettings.FirstOrDefault(p => p.ID == id);
            if (fieldsetting != null)
            {
                var fieldsettingView = (FieldSettingView)ModelMapper.Map(fieldsetting, typeof(Settings), typeof(FieldSettingView));
                return View(fieldsettingView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditFieldSetting(FieldSettingView fieldView)
        {
            if (ModelState.IsValid)
            {
                var fieldsetting = Repository.FieldSettings.FirstOrDefault(p => p.ID == fieldView.ID);
                ModelMapper.Map(fieldView, fieldsetting, typeof(FieldSettingView), typeof(Settings));
                Repository.UpdateFieldSetting(fieldsetting);

                return RedirectToAction("Index");
            }

            return View(fieldView);
        }


        [HttpGet]
        public ActionResult DeleteFieldSetting(int? id)
        {
            var field = Repository.FieldSettings.FirstOrDefault(p => p.ID == id);
            if (field != null)
            {
                Repository.RemoveFieldSetting(field);
            }
            return RedirectToAction("Index");
        }
    }
}
