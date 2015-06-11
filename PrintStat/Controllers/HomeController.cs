using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using Ninject;
using PrintStat.Models;
using PrintStat.Models.ViewModels;


namespace PrintStat.Controllers
{
    public class HomeController : BaseController
    {

        public ActionResult GetMail()
        {
           // (new ParseModule.GetData(new DateTime(2015,06,01),DateTime.Now)).GetSource();
            //return View();
            return null;
        }

        public ActionResult Index(int page=1)
        {
           // var Jobs = Repository.Jobs.OrderByDescending(s => s.EndTime).ToList();
            //var Jobs = Repository.JobPaginators(page).ToList();

            ViewBag.pageCount = null;
            
            return View();
        }

        public ActionResult Index1()
        {
            
            return View();
        }



        private void InitViewBag()
        {
            ViewBag.Printers = Repository.Printers;
            ViewBag.Plotters = Repository.Plotters;
            ViewBag.Applications = Repository.Applications;
            ViewBag.PrinterPaper = Repository.PrinterPaperTypes;
            ViewBag.SizePaper = Repository.SizePapers;
            ViewBag.PlotterPapertypes = Repository.PlotterPaperTypes;
            ViewBag.AuthorEmployees = Repository.AuthorEmployees;
            ViewBag.UserEmployees = Repository.UserEmployees;
            ViewBag.PaperTypes = Repository.PaperTypes;
        }

        [AcceptVerbs(HttpVerbs.Get)]
        public JsonResult LoadSizePaperForDevice(int idDev)
        {
            var sizePapers = Repository.Printers.First(p => p.ID == idDev).Model.SupportSize
                .Join(Repository.SizePapers,p=>p.SizePaperID,s=>s.ID,(p,s)=> new{
                ID= s.ID,
                Name = s.Name
                });
            var sizePaperData = sizePapers.Select(m => new SelectListItem()
            {
                Text = m.Name,
                Value = m.ID.ToString()
            });
            return Json(sizePaperData, JsonRequestBehavior.AllowGet);
        }        [AcceptVerbs(HttpVerbs.Get)]
        public JsonResult LoadPaperTypeForDevice(int idDev)
        {
            var paperTypes = Repository.Printers.First(p => p.ID == idDev).Model.ModelPaperType
                .Join(Repository.PaperTypes,p=>p.PaperTypeID,s=>s.ID,(p,s)=> new{
                ID= s.ID,
                Name = s.Name
                });
            var paperTypeData = paperTypes.Select(m => new SelectListItem()
            {
                Text = m.Name,
                Value = m.ID.ToString()
            });
            return Json(paperTypeData, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult CreatePrinterJob()
        {

            InitViewBag();
            var newJobView = new JobView();
            newJobView.StartTime = DateTime.Now;
            newJobView.EndTime = DateTime.Now;
            newJobView.Duration = 0;
            newJobView.Pages = 1;
            newJobView.Copies = 1;
            newJobView.SizePaperID = 0;
            newJobView.PaperTypeID = 0;
            newJobView.IsManual = true;

            return View(newJobView);
        }

        public ActionResult CreatePlotterJob()
        {

            InitViewBag();
            var newJobView = new JobView();
            newJobView.StartTime = DateTime.Now;
            newJobView.EndTime = DateTime.Now;
            newJobView.Duration = 0;
            newJobView.Pages = 1;
            newJobView.Copies = 1;
            newJobView.SizePaperID = 0;
            newJobView.PaperTypeID = 0;
            newJobView.IsManual = true;

            return View(newJobView);
        }

        [HttpPost]
        public ActionResult CreateJob(JobView JobView)
        {
            var anyJob = Repository.Jobs.Any(p => string.Compare(p.Name, JobView.Name) == 0);
            if (anyJob)
            {
                ModelState.AddModelError("Name", "Задание с таким наименованием уже существует");
            }

            if (ModelState.IsValid)
            {

                var Job = (Job)ModelMapper.Map(JobView, typeof(JobView), typeof(Job));
                Repository.CreateJob(Job);
                return RedirectToAction("Index");
            }

            return View(JobView);
        }


        [HttpGet]
        public ActionResult EditPrinterJob(int? id)
        {


            var Job = Repository.Jobs.FirstOrDefault(p => p.ID == id);
            if (Job != null)
            {
                InitViewBag();
                var JobView = (JobView)ModelMapper.Map(Job, typeof(Job), typeof(JobView));
                return View(JobView);
            }
            return RedirectToAction("Index");
        }

        [HttpGet]
        public ActionResult EditPlotterJob(int? id)
        {


            var Job = Repository.Jobs.FirstOrDefault(p => p.ID == id);
            if (Job != null)
            {
                InitViewBag();
                var JobView = (JobView)ModelMapper.Map(Job, typeof(Job), typeof(JobView));
                return View(JobView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditJob(JobView JobView)
        {
            InitViewBag();
            if (ModelState.IsValid)
            {
                var Job = Repository.Jobs.FirstOrDefault(p => p.ID == JobView.ID);
                ModelMapper.Map(JobView, Job, typeof(JobView), typeof(Job));
                Repository.UpdateJob(Job);

                return RedirectToAction("Index");
            }

            return View(JobView);
        }


        [HttpGet]
        public ActionResult DeleteJob(int? id)
        {
            var Job = Repository.Jobs.FirstOrDefault(p => p.ID == id);
            if (Job != null)
            {
                Repository.RemoveJob(Job);
            }
            return RedirectToAction("Index");
        }
    

        public ActionResult About()
        {
            ViewBag.Message = "Your app description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public class temp
        {
            public string Name { get; set; }
            public int Value { get; set; }
        } 
        
        public IEnumerable<temp> GetDevice()
        {

            var result= new List<temp>();

            result.Add(new temp {Name = "Принтер", Value = Repository.DeviceTypes.Where(p=>p.Name=="Принтер").Join(Repository.Models,m=>m.ID,dt=>dt.DeviceTypeID,
                (dt,m)=>m).Join(Repository.PrintersAndPlotters,m=>m.ID,d=>d.ModelID,
                (m,d)=>d).Count()});
            result.Add(new temp
            {
                Name = "Плоттер",Value = Repository.DeviceTypes.Where(p => p.Name == "Плоттер").Join(Repository.Models, m => m.ID, dt => dt.DeviceTypeID,
                    (dt, m) => m).Join(Repository.PrintersAndPlotters, m => m.ID, d => d.ModelID,
                    (m, d) => d).Count()
            });


            return result;
        }

        public IEnumerable<temp> GetJobsDevice()
        {
            var result = new List<temp>();
            
            foreach (var dev in Repository.PrintersAndPlotters)
            {
                result.Add(new temp{Name = dev.Name, Value = dev.Job.Count});
            }
            //return Json(new { DevsJob = result }, JsonRequestBehavior.AllowGet);
            return result;
        }

        public ActionResult GetChart1()
        {
            var data = GetDevice();
            var myChart = new Chart(width: 600, height: 400)
                .AddTitle("Соотношение принтеров и плоттеров")
                .AddSeries(
                    name: "Соотношение принтеров и плоттеров",
                    
                    chartType: "Pie",
                    xValue: data, xField: "Name",
                    yValues: data, yFields: "Value")
                .AddLegend("","Name")
                .Write();
            return null;
        }     
        public ActionResult GetChart()
        {
            var data = GetJobsDevice();
            var myChart = new Chart(width: 600, height: 400, theme: ChartTheme.Yellow)

            .AddTitle("Количество заданий печати на каждое устройство печати")
            .DataBindTable(dataSource: data)

            .AddSeries("Default",
                xValue: data, xField: "Name",
                yValues: data, yFields: "Value")

            .Write();
            return null;
        }
    }
}
