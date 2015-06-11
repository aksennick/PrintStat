using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;


namespace PrintStat.Models.ViewModels
{
    public class BaseView
    {
        public int ID { get; set; }
        [Display (Name = "Наименование")]
        [Required(ErrorMessage = "Введите наименование")]

        public string Name { get; set; }
    }
}