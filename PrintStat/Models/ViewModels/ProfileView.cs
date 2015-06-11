using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PrintStat.Models.ViewModels
{
    public class ProfileView : BaseView
    {
        //
        // GET: /CartridgeColorView/

        public class setval
        {
            public int idValue { get; set; }

            [Display(Name = "Значение")]
            public string value { get; set; }
            public string setting { get; set; }
        }

       
        private List<setval> _setval = new List<setval>();
        public List<setval> SettingVals
        {
            get { return _setval; }
            set { _setval = value; }
        }
        public IQueryable<Settings> Setting { get; set; }
        public int[] ChosenSetting { get; set; }
    }
}