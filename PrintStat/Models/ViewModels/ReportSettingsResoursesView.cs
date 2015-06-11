using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PrintStat.Models.ViewModels
{
    public class ReportSettingsResoursesView
    {
        public int[] ChosenRules { get; set; }

        //private List<Model> _Models = new List<Model>();
        //public List<Model> Models
        //{
        //    get { return _Models; }
        //    set { _Models = value; }
        //}

    }
}