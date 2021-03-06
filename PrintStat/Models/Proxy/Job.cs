﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PrintStat
{
    public partial class Job: IBaseObject
    {


        public override string ToString()
        {
            return Name;
        }

        public Department Department
        {
            get 
            {
                if (Author != null) return Author.Department;
                return null;
            }
        }

        public string ComputerNameOrIP
        {
            get
            {
                var s = "";

                if (ComputerName != null) s = ComputerName;
                if (s != "") s = string.Concat(s, "/");
                if (IP != null) s = string.Concat(s, IP);

                return s;

            }
        }

        public DeviceType DeviceType
        {
            get
            {
                return Printer.DeviceType;
            }
        }

        public decimal RealWidth_cm
        {
            get
            {
                if (_Width_cm == 0 && PaperType != null && PaperType.Width_cm != null) return PaperType.Width_cm.Value;
                return _Width_cm;
            }
        }

        public decimal RealHeight_cm
        {
            get
            {
                if (_Height_cm == 0 && PaperType != null && PaperType.Height_cm != null) return PaperType.Height_cm.Value;
                return _Width_cm;
            }
        }

        public string PrinterName
        {
            get { return Printer != null?Printer.Name:""; }
        }

        public string AuthorName
        {
            get {  return Author != null?Author.Name:""; }
        }

        public string EmployeeName
        {
            get { return Employee!= null ? Employee.Name : ""; }
        }

        public string DepartmentName
        {
            get
            {
                return Department != null ? Department.Name : "";
            }
        }

        public string PaperTypeName
        {
            get
            {
                return PaperType != null ? PaperType.Name : "";
            }
        }

        public string ApplicationName
        {
            get
            {
                return Application != null ? Application.Name : "";
            }
        }
    }
}