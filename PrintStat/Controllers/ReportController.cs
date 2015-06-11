using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;


namespace PrintStat.Controllers
{
    public class ReportController : BaseController
    {
        //
        // GET: /Report/

        class Rule
        {
            public int i;
            public string value;
            public string text;
        }
        public JsonResult _GetRule()
        {

            var rules = new List<Rule>();
            rules.Add(new Rule(){ text = "Устройство",value = "1"});
            rules.Add(new Rule(){ text ="Исполнитель",value = "2"});
            rules.Add(new Rule(){ text="Автор задания", value = "3"});
            rules.Add(new Rule(){ text= "Отдел", value = "4" });


            return Json(rules, JsonRequestBehavior.AllowGet);



        }
        public IEnumerable<SelectListItem> PrintersAndPlotters
        {
            get
            {
                yield return new SelectListItem {Value = "0", Text = "Не задан"};
                foreach (var p in Repository.PrintersAndPlotters)
                {
                    yield return new SelectListItem { Value = p.ID.ToString(), Text = p.Name };
                }
                
            }
        }

        public IEnumerable<SelectListItem> Applications
        {
            get
            {
                yield return new SelectListItem { Value = "0", Text = "Не задано" };
                foreach (var p in Repository.Applications)
                {
                    yield return new SelectListItem { Value = p.ID.ToString(), Text = p.Name };
                }

            }
        }

        public IEnumerable<SelectListItem> AuthorEmployees
        {
            get
            {
                yield return new SelectListItem { Value = "", Text = "Не задан" };
                foreach (var p in Repository.AuthorEmployees)
                {
                    yield return new SelectListItem { Value = p.TabNumber, Text = p.Name };
                }

            }
        }


        public IEnumerable<SelectListItem> UserEmployees
        {
            get
            {
                yield return new SelectListItem { Value = "", Text = "Не задан" };
                foreach (var p in Repository.UserEmployees)
                {
                    yield return new SelectListItem { Value = p.TabNumber, Text = p.Name };
                }

            }
        }

        public IEnumerable<SelectListItem> Departments
        {
            get
            {
                yield return new SelectListItem { Value = "0", Text = "Не задан" };
                foreach (var p in Repository.Departments)
                {
                    yield return new SelectListItem { Value = p.ID.ToString(), Text = p.Name };
                }

            }
        }

        public ActionResult ExportJobs()
        {
            ViewBag.Printers = PrintersAndPlotters;
            ViewBag.Applications = Applications;
            ViewBag.AuthorEmployees = AuthorEmployees;
            ViewBag.UserEmployees =  UserEmployees;
            ViewBag.Departments = Departments;
            var reportView = new ReportSettingsView()
            {
                DateStart = DateTime.Now,
                DateEnd = DateTime.Now,
            };
            return View(reportView);
        }

        [HttpPost]
        public ActionResult ExportJobs(ReportSettingsView reportsettingsview)
        {
            GridView gv = new GridView();


            var subjobs = Repository.Jobs.Where(p => ((p.StartTime == reportsettingsview.DateStart && p.EndTime == reportsettingsview.DateEnd) &&
                p.ApplicationID == reportsettingsview.ApplicationID || reportsettingsview.ApplicationID == 0) &&
                ((p.Author != null && p.Author.TabNumber == reportsettingsview.AuthorTabNumber) || reportsettingsview.AuthorTabNumber == null) &&
                                                   ((p.Department != null && p.Department.ID == reportsettingsview.DepartmentID) || reportsettingsview.DepartmentID == 0) &&
                 ((p.Employee != null && p.Employee.TabNumber == reportsettingsview.EmployeeTabNumber) || reportsettingsview.EmployeeTabNumber == null) &&
                                                   ((p.Device != null && p.Device.ID == reportsettingsview.PrinterID) || reportsettingsview.PrinterID == 0)).OrderByDescending(s => s.EndTime);

          
            var jobs = from j in subjobs
                       select new { j.Name, j.Pages, j.Copies, j.PrinterName, j.ApplicationName, j.RealWidth_cm, j.RealHeight_cm, j.PaperTypeName, j.StartTime, j.EndTime, j.Duration, j.EmployeeName, j.AuthorName, j.DepartmentName };


        try
            {
            gv.DataSource = jobs;
            gv.DataBind();
            
            gv.HeaderRow.Cells[0].Text = "Название задания";
            gv.HeaderRow.Cells[1].Text = "Страниц";
            gv.HeaderRow.Cells[2].Text = "Копий";
            gv.HeaderRow.Cells[3].Text = "Устройство";
            gv.HeaderRow.Cells[4].Text = "Приложение";
            gv.HeaderRow.Cells[5].Text = "Ширина, см";
            gv.HeaderRow.Cells[6].Text = "Длина, см";
            gv.HeaderRow.Cells[7].Text = "Типоразмер";
            gv.HeaderRow.Cells[8].Text = "Начало печати";
            gv.HeaderRow.Cells[9].Text = "Окончание печати";
            gv.HeaderRow.Cells[10].Text = "Длительность, мин";
            gv.HeaderRow.Cells[11].Text = "Отправитель";
            gv.HeaderRow.Cells[12].Text = "Автор";
            gv.HeaderRow.Cells[13].Text = "Отдел";

            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=PrintJobReport.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);
            gv.RenderControl(htw);
            Response.Output.Write(sw.ToString());
            Response.Flush();
            Response.End();
                return RedirectToAction("Index");
            }
            catch (Exception)
            {

                ViewBag.Message = "Результатов не найдено, попробуйте изменить критерии поиска ";
                return View("ErrorView");
            }


            
        }

        [HttpGet]
        public ActionResult ExportResourses()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ExportResourses(ReportSettingsResoursesView viewModel)
        {
            //todo основная часть запроса
            foreach (var item in viewModel.ChosenRules)
            {
                
                // если выбраны критерии
                 switch (item)
                {
                     case 1:
                        break;
                     case 2:
                        break;
                     case 3:
                        break;
                     case 4:
                        break;

                };
                //формируем exel
                try
                {

                }
                catch (Exception)
                {
                    
                    throw;
                }
            }
           
            return RedirectToAction("Index");
        }
    }
}
