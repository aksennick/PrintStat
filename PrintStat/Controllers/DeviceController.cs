using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class DeviceController : BaseController
    {
        //



        public ActionResult Index()
        {
                var printers = Repository.PrintersAndPlotters.ToList();
                return View(printers);
        }


        private void 
        InitViewBag()
        {
            ViewBag.Models = Repository.Models;
        }
        [HttpGet]
        public ActionResult CreateDevice()
        {
            InitViewBag();
            var newPrintView = new DeviceView ();
            return View(newPrintView);
        }



        [HttpPost]
        public ActionResult CreateDevice(DeviceView  printerView)
        {
            InitViewBag();
            var anyPrinter = Repository.PrintersAndPlotters.Any(p => string.Compare(p.Name, printerView.Name)==0);  
            if (anyPrinter)
            {
                ModelState.AddModelError("Name", "Принтер с таким наименованием уже существует");
            }
            if (printerView.ModelID == 0)
            {
                ModelState.AddModelError("ModelID","Выберите модель");
            }

            if (ModelState.IsValid)
            {
                

                var printer = (Device)ModelMapper.Map(printerView, typeof(DeviceView), typeof(Device));
                Repository.CreatePrinter(printer);
                var idDevice = printer.ID;
                var consumableModel = Repository.GetModelConsumables(idDevice);
                foreach (var item in consumableModel)
                {
                    Repository.CreateDeviceConsumable(idDevice,item.ID,DateTime.MinValue);
                }

                return RedirectToAction("Index");
            }

            return View(printerView);
        }


        [HttpGet]
        public ActionResult EditDevice(int? id)
        {
            InitViewBag();
            var printer = Repository.PrintersAndPlotters.FirstOrDefault(p => p.ID == id);
            if (printer != null)
            {
                var printerView = (DeviceView)ModelMapper.Map(printer, typeof(Device), typeof(DeviceView));
                return View(printerView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditDevice(DeviceView printerView)
        {
            if (ModelState.IsValid)
            {
                var printer = Repository.PrintersAndPlotters.FirstOrDefault(p => p.ID == printerView.ID);
                ModelMapper.Map(printerView, printer, typeof(DeviceView), typeof(Device));
                Repository.UpdatePrinter(printer);

                return RedirectToAction("Index");
            }

            return View(printerView);
        }


        [HttpGet]
        public ActionResult DeleteDevice(int? id)
        {
            var printer = Repository.PrintersAndPlotters.FirstOrDefault(p => p.ID == id);
            if (printer != null)
            {
                Repository.RemoveDeviceConsumable((int)id);
                Repository.RemovePrinter(printer);
            }
            return RedirectToAction("Index");
        }



        [HttpGet]
        public ActionResult EditDeviceConsumable(int? id)
        {
            var device = Repository.PrintersAndPlotters.FirstOrDefault(p => p.ID == id);
            if (device != null)
            {
                ViewBag.DeviceName = device.Name;
                var deviceConsumble = new DeviceConsumableView();
                deviceConsumble.deviceID = (int) id;
                deviceConsumble.AllDevCons = Repository.GetDevConses(id);
                return View(deviceConsumble);


            }
            return View("Index");
        }



        //option 1 прекратить использование
        //       2 заменить
        [HttpGet]
#pragma warning disable 1998
        public async Task<ActionResult> GetDeviceConsumble(int deviceId, int id, int option )
#pragma warning restore 1998
        {
            if (option == 1)
            {
                Repository.SetDateEnd(id, DateTime.MinValue);
            }
            else
            {
                Repository.SetUseOffAndAddNewCons(id);
            }

            var deviceConsumble = new DeviceConsumableView
            {
                AllDevCons = Repository.GetDevConses(deviceId),
                deviceID = deviceId
            };

            return PartialView("partialDeviceConsumable", deviceConsumble);
        }


    }

}
