using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using Microsoft.SqlServer.Server;
using PrintStat.Models.ViewModels;


namespace PrintStat.Models
{
    public partial class SQLRepository : IRepository
    {

        public IQueryable<Profile> Profiles
        {
            get
            {
                return Db.Profile;
            }
        }

        public bool CreateProfile(Profile instance)
        {
            if (instance.ID == 0)
            {
                Db.Profile.InsertOnSubmit(instance);
                Db.Profile.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool UpdateProfile(Profile instance)
        {
            Profile cache = Db.Profile.Where(p => p.ID == instance.ID).FirstOrDefault();
            if (cache != null)
            {
                //TODO : Update fields for Profile
                Db.Profile.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool RemoveProfile(Profile instance)
        {
            if (instance != null)
            {
                Db.Profile.DeleteOnSubmit(instance);
                Db.Profile.Context.SubmitChanges();
                return true;
            }

            return false;
        }




        public IQueryable<Settings> Settingses
        {
            get
            {
                return Db.Settings;
            }
        }

        public bool CreateSettings(Settings instance)
        {
            if (instance.ID == 0)
            {
                Db.Settings.InsertOnSubmit(instance);
                Db.Settings.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool UpdateSettings(Settings instance)
        {
            Settings cache = Db.Settings.Where(p => p.ID == instance.ID).FirstOrDefault();
            if (cache != null)
            {
                //TODO : Update fields for Settings
                Db.Settings.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool RemoveSettings(Settings instance)
        {

            if (instance != null)
            {
                Db.Settings.DeleteOnSubmit(instance);
                Db.Settings.Context.SubmitChanges();
                return true;
            }

            return false;
        }



        public IQueryable<SettingValue> SettingValues
        {
            get
            {
                return Db.SettingValue;
            }
        }

        public bool CreateSettingValue(int[] setId, int profId)
        {
            try
            {

                Db.SettingValue.InsertAllOnSubmit(
                   setId.Select(id =>
                       new SettingValue()
                       {
                           SettingsID = id,
                           ProfileID = profId
                       })
                   );
               Db.SettingValue.Context.SubmitChanges();
                        return true;
            }
            catch (Exception)
            {
                return false;
                throw;
            }

            
        }

        public bool UpdateSettingValue(List<ProfileView.setval> sv)
        {
           // SettingValue cache1 = Db.SettingValue.Where(p => p.ProfileID == idProfile &&).FirstOrDefault();

            foreach (var item in sv)
            {
                SettingValue cache = Db.SettingValue.First(p => p.ID == item.idValue);
                cache.Value = item.value;
                Db.SettingValue.Context.SubmitChanges();
                
            }

            return true;
        }

        public bool RemoveSettingValue(SettingValue instance)
        {
            if (instance != null)
            {
                Db.SettingValue.DeleteOnSubmit(instance);
                Db.SettingValue.Context.SubmitChanges();
                return true;
            }

            return false;
        }
        
        
    }
}