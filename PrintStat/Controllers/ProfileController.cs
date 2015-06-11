using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class ProfileController : BaseController
    {
        //
        // GET: /Profile/
        public IEnumerable<SelectListItem> Protocols
        {
            get
            {
                var p = new List<SelectListItem>();
                //p.Add(new SelectListItem() { Value = "", Text = "Выберите протокол" });
                p.Add(new SelectListItem() { Value = "POP3", Text = "POP3" });
                p.Add(new SelectListItem() { Value = "IMAP", Text = "IMAP" });
                return p;
            }
        }

        public IEnumerable<SelectListItem> Names
        {
            get
            {
                var p = new List<SelectListItem>();
                p.Add(new SelectListItem() {Value = "Email", Text = "Email"});
                p.Add(new SelectListItem() { Value = "ActiveDirectory", Text = "MS ActiveDirectory" });
                return p;
            }
        }



        public ActionResult Index()
        {
            var profiles = Repository.Profiles.ToList();
            return View(profiles);
        }




        [HttpGet]
        public ActionResult CreateProfile()
        {
            ViewBag.Names = Names;
            var profileView = new ProfileView {Setting = Repository.Settingses.AsQueryable()};
            return View(profileView);
        }

        [HttpPost]
        public ActionResult CreateProfile(ProfileView profileView)
        {
            var anyProfile = Repository.Profiles.Any(p => String.Compare(p.Name, profileView.Name) == 0);
            if (anyProfile)
            {
                ModelState.AddModelError("Name","Профиль с таким название уже существует");
            }
            else
            {
                var temp = new Profile() {Name = profileView.Name};
                Repository.CreateProfile(temp);
            
                if (ModelState.IsValid)
                {
                    Repository.CreateSettingValue(profileView.ChosenSetting,temp.ID);
                    return RedirectToAction("Index");
                }
            }
            ViewBag.Names = Names;
            profileView.Setting = Repository.Settingses.AsQueryable();
            return View(profileView);
        }
        [HttpGet]
        public ActionResult EditProfile(int? id)
        {
            ViewBag.Protocols = Protocols;
            
            var profil= Repository.Profiles.FirstOrDefault(p => p.ID == id);
            if (profil != null)
            {
                ViewBag.Title = profil.Name;
                var profileView = (ProfileView) ModelMapper.Map(profil, typeof (Profile), typeof (ProfileView));


                var setingVal = Repository.Profiles.Where(p => p.ID == id).Join(
                    Repository.SettingValues,
                    prof => prof.ID,
                    setv => setv.ProfileID,
                    (prof, setv) => setv).Join(
                        Repository.Settingses,
                        setv => setv.SettingsID,
                        set => set.ID,
                        (setv, set) => new {idVal = setv.ID, Val = setv.Value, setting = set.Name});

                List<ProfileView.setval> temp = new List<ProfileView.setval>();
                foreach (var p in setingVal)
                {
                    var sv = new ProfileView.setval();
                    sv.idValue = p.idVal;
                    sv.value = p.Val;
                    sv.setting = p.setting;
                    temp.Add(sv);
                }
                profileView.SettingVals = temp;
                return View(profileView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditProfile(ProfileView profileView)
        {
            ViewBag.Protocols = Protocols;
            {
                Repository.UpdateSettingValue(profileView.SettingVals);
                return RedirectToAction("Index");    
            }
           
        }     
        [HttpGet]
        public ActionResult DeleteProfile(int? id)
        {
            var temp = Repository.Profiles.First(p => p.ID == id);
            {
                Repository.RemoveProfile(temp);
                return RedirectToAction("Index");    
            }
           
        }
    }
}
