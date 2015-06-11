using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
using PrintStat.Models.Validation;
using DataAnnotationsExtensions;


namespace PrintStat.Models.ViewModels
{
    public class JobView: BaseView
    {
        [Display(Name = "Устройство")]
        [Required(ErrorMessage = "Укажите устройство")]
        public int DeviceID { get; set; }
        [Display(Name = "Приложение")]
        [Required(ErrorMessage = "Укажите приложение")]
        public int ApplicationID { get; set; }
        public int? Duration { get; set; }
        [Display(Name = "Начало выполнения задания")]
        public DateTime? StartTime { get; set; }
        [Display(Name = "Окончание выполнения задания")]
        public DateTime? EndTime { get; set; }
        [Display(Name = "Исполнитель")]
        [Required(ErrorMessage = "Укажите исполнителя")]
        public string UserTabNumber { get; set; }
        [IsNumeric]
        [Display(Name = "Количество страниц")]
        public int Pages { get; set; }
        [IsNumeric]
        [Display(Name = "Количество копий")]
        public int Copies { get; set; }
        [IsNumeric]
        [Display(Name = "Ширина")]
        [Required(ErrorMessage = "Укажите ширину")]
        public decimal Width_cm { get; set; }
        [IsNumeric]
        [Display(Name = "Длина")]
        [Required(ErrorMessage = "Укажите длину")]
        public decimal Height_cm { get; set; }
        [IsNumeric]
        [Display(Name = "Ширина в пикселях")]

        public int? Width_px { get; set; }
        [IsNumeric]
         [Display(Name = "Длина в пикселях")]
        public int? Height_px { get; set; }
         [Display(Name = "Типоразмер бумаги")]
        public int? SizePaperID { get; set; }
         [Display(Name = "Тип бумаги")]
        public int? PaperTypeID { get; set; }
         [Display(Name = "Автор работы")]
        public string AuthorTabNumber { get; set; }
         [Display(Name = "Размер файла")]
        public int? Size_kb { get; set; }
         [Display(Name = "IP адрес")]
        public string IP { get; set; }
         [Display(Name = "Доменное имя компьютера")]
        public string ComputerName { get; set; }
         [Display(Name = "Ручной ввод")]
        public bool IsManual { get; set; }



    }
}