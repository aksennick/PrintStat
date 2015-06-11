using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PrintStat.Models.ViewModels
{
    public class ReportSettingsView 
    {
        //time
        [DataType(DataType.Date)]
        public DateTime DateStart { get; set; }
        [DataType(DataType.Date)]
        public DateTime DateEnd { get; set; }



        public string EmployeeTabNumber { get; set; }
        public int ApplicationID { get; set; }
        public string AuthorTabNumber { get; set; }
        public int DepartmentID { get; set; }
        public int PrinterID { get; set; }

    }
}