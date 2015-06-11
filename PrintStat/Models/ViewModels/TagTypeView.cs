using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace PrintStat.Models.ViewModels
{
    public class TagTypeView
    {
        public int ID { get; set; }
        [Display(Name = "Тип тега")]
        [Required(ErrorMessage = "Введите тип тега")]
        [StringLength(255)]
        public string Type { get; set; }
    }
}