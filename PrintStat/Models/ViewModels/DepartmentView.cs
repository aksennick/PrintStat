using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PrintStat.Models.ViewModels
{
    public class DepartmentView : BaseView
    {
        //
        // GET: /DepartmentView/
        [Display(Name = "Короткое название")]
        [Required(ErrorMessage = "Введите короткое название")]
        [StringLength(20)]
        public string ShortName { get; set; }
    }
}
