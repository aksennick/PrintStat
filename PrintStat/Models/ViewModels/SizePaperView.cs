using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
using PrintStat.Models.Validation;
using DataAnnotationsExtensions;


namespace PrintStat.Models.ViewModels
{
    public class SizePaperView: BaseView
    {

        [IsNumeric]
        [Display(Name = "Ширина, см")]
        public decimal? Width_cm { get; set; }

        [IsNumeric]
        [Display(Name = "Длина, см")]
        public decimal? Height_cm { get; set; }
    }
}