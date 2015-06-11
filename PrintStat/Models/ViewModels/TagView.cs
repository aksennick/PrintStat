using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace PrintStat.Models.ViewModels
{
    public class TagView :BaseView
    {
        [Display(Name = "Тег")]
        [Required(ErrorMessage = "Введите тег")]
        [StringLength(30)]
        public string Tag1 { get; set; }
        [Display(Name = "Тип тега")]
        [Required(ErrorMessage = "Укажите тип тега")]
        [StringLength(30)]
        public int TagTypeID { get; set; }

    }
}