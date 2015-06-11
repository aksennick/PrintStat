using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class EmployeeController : BaseController
    {
        //
        // GET: /Employee/

        public ActionResult Index()
        {
            var employee = Repository.Employees.ToList();
            return View(employee);
        }

        [HttpGet]
        public ActionResult CreateEmployee()
        {
            var newemployeeView = new EmployeeView();
            return View(newemployeeView);
        }

        [HttpPost]
        public ActionResult CreateEmployee(EmployeeView employeeView)
        {
            var anyemployee = Repository.Employees.Any(p => string.Compare(p.Name, employeeView.Name) == 0);
            if (anyemployee)
            {
                ModelState.AddModelError("Name", "Такой сотрудник уже существует");
            }

            if (ModelState.IsValid)
            {

                var employee = (Employee)ModelMapper.Map(employeeView, typeof(EmployeeView), typeof(Employee));
                Repository.CreateEmployee(employee);
                return RedirectToAction("Index");
            }

            return View(employeeView);
        }

        [HttpGet]
        public ActionResult EditEmployee(int? id)
        {
            var employee = Repository.Employees.FirstOrDefault(p => p.ID == id);
            if (employee != null)
            {
                var employeeView = (EmployeeView)ModelMapper.Map(employee, typeof(Employee), typeof(EmployeeView));
                return View(employeeView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditEmployee(EmployeeView employeeView)
        {
            if (ModelState.IsValid)
            {
                var employee = Repository.Employees.FirstOrDefault(p => p.ID == employeeView.ID);
                ModelMapper.Map(employeeView, employee, typeof(EmployeeView), typeof(Employee));
                Repository.UpdateEmployee(employee);

                return RedirectToAction("Index");
            }

            return View(employeeView);
        }


        [HttpGet]
        public ActionResult DeleteEmployee(int? id)
        {
            var employee = Repository.Employees.FirstOrDefault(p => p.ID == id);
            if (employee != null)
            {
                Repository.RemoveEmployee(employee);
            }
            return RedirectToAction("Index");
        }

    }
}
