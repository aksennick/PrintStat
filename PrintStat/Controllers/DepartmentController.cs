using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
    public class DepartmentController : BaseController
    {
        //
        // GET: /Department/

        public ActionResult Index()
        {
            var departments = Repository.Departments.ToList();
            return View(departments);
        }

        [HttpGet]
        public ActionResult CreateDepartment()
        {
            var departmentsView = new DepartmentView();
            return View(departmentsView);
        }

        [HttpPost]
        public ActionResult CreateDepartment(DepartmentView departmentView)
        {
            var anyDepartment = Repository.Departments.Any(p => string.Compare(p.Name, departmentView.Name) == 0);
            if (anyDepartment)
            {
                ModelState.AddModelError("Name", "Отдел с таким наименованием уже существует");
            }

            if (ModelState.IsValid)
            {

                var department = (Department)ModelMapper.Map(departmentView, typeof(DepartmentView), typeof(Department));
                Repository.CreateDepartment(department);
                return RedirectToAction("Index");
            }

            return View(departmentView);
        }

        [HttpGet]
        public ActionResult EditDepartment(int? id)
        {

            var department = Repository.Departments.FirstOrDefault(p => p.ID == id);
            if (department != null)
            {
                var departmentView = (DepartmentView)ModelMapper.Map(department, typeof(Department), typeof(DepartmentView));
                return View(departmentView);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditDepartment(DepartmentView departmentView)
        {
            var anyDepartment = Repository.Departments.Where(p=>p.ID!=departmentView.ID).Any(p => string.Compare(p.Name, departmentView.Name) == 0);
            if (anyDepartment)
            {
                ModelState.AddModelError("Name", "Отдел с таким наименованием уже существует");
            }
            if (ModelState.IsValid)
            {
                var department = Repository.Departments.FirstOrDefault(p => p.ID == departmentView.ID);
                ModelMapper.Map(departmentView, department, typeof(DepartmentView), typeof(Department));
                Repository.UpdateDepartment(department);

                return RedirectToAction("Index");
            }

            return View(departmentView);
        }


        [HttpGet]
        public ActionResult DeleteDepartment(int? id)
        {
            var department = Repository.Departments.FirstOrDefault(p => p.ID == id);
            if (department != null)
            {
                if (!Repository.RemoveDepartment(department))
                {
                    ViewBag.Message = "Невозможно удалить значение, т.к. оно используется";
                    return View("~/Views/Shared/ErrorView.cshtml");
                }
            }
            return RedirectToAction("Index");
        }

    }
}
