using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PrintStat.Models.ViewModels
{
    public class CartridgeColorView : BaseView
    {
        //
        // GET: /CartridgeColorView/

        [Display(Name = "Короткое название")]
        [Required(ErrorMessage = "Введите короткое название")]
        [StringLength(2)]
        public string ShortName { get; set; }

    }
}
