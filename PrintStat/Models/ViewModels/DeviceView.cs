using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PrintStat.Models.ViewModels
{
    public class DeviceView: BaseView
    {
        private List<Model> _model = new List<Model>();

        public List<Model> models
        {
            get { return _model; }
        }
        [Display(Name = "Производитель")]
        [Required(ErrorMessage = "Пожалуйста введите производителя")]
        public string Manufacturer { get; set; }
        [Display(Name = "Модель")]
        [Required(ErrorMessage = "Пожалуйста выберите модель")]
        public int ModelID { get; set; }
        [Display(Name = "Серийный номер")]
        [Required(ErrorMessage = "Пожалуйста введите серийный номер")]
        public string Sn { get; set; }
        [Display(Name = "Поисковая строка")]
        [Required(ErrorMessage = "Пожалуйста введите поисковую строку")]
        public string SearchString { get; set; }
        [Display(Name = "Инвентарный номер")]
        public string InvNumber { get; set; }
        [Display(Name = "Поддержка отправки статистики")]
        public bool StatisticsSupported { get; set; }
    }
}