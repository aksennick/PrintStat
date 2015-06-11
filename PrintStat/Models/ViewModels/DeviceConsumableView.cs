using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DotNetOpenAuth.OpenId.Extensions.AttributeExchange;

namespace PrintStat.Models.ViewModels
{
    public class DeviceConsumableView : BaseView
    {
        public List<SQLRepository.devCons> AllDevCons { set; get; }

        [Required(ErrorMessage = "Укажите устройство")]
        public int deviceID { get; set; }

       // [Display(Name = "Рабочий ресурс")]
       // public int Endurance { get; set; }

      //  [Display(Name = "Короткое название")]
      //  public string ShortName { get; set; }

       // [Display(Name = "Тип комплектующего")]
       // [Required(ErrorMessage = "Выберите тип комплектующего")]
       // public int TypeConsumableID { get; set; }
        [Display(Name = "Поддерживаемые модели")]
        public int[] ChosenModelIds { get; set; }

        private List<Model> _Models = new List<Model>();
        public List<Model> Models
        {
            get { return _Models; }
            set { _Models = value; }
        }
    }
}
