using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;


namespace PrintStat.Controllers
{
    public class SettingsController : BaseController
    {
        //
        // GET: /Settings/

        public IEnumerable<SelectListItem> Protocols
        {
            get
            {
                var p = new List<SelectListItem>();
                p.Add(new SelectListItem() { Value = "POP3", Text = "POP3" });
                p.Add(new SelectListItem() { Value = "IMAP", Text = "IMAP" });
                return p;
            }
        }

        //public ActionResult Index()
        //{

        //    var settings = Repository.Setup.FirstOrDefault();

        //    SettingsView settingsView;
        //    if (settings != null)
        //        settingsView = (SettingsView)ModelMapper.Map(settings, typeof(Setup), typeof(SettingsView));
        //    else
        //        settingsView =  new SettingsView();

        //    ViewBag.Protocols = Protocols;  
        //    return View(settingsView);
        //}


      //  [HttpPost]
        //public ActionResult SaveSettings(SettingsView settingsView)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        var settings = (Setup)ModelMapper.Map(settingsView, typeof(SettingsView), typeof(Setup));
        //        Repository.SaveSettings(settings);
        //        return RedirectToAction("Index");
        //    }
        //    return View(settingsView);
        //}
    }
}
