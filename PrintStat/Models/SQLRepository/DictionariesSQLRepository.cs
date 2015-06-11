using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using Ninject.Infrastructure.Language;


namespace PrintStat.Models
{
    public partial class SQLRepository : IRepository
    {

        #region typeComponenet
        [Display(Name = "Тип комрдектующего")]
        public IQueryable<TypeConsumable> TypeConsumables
        {
            get
            {
                return Db.TypeConsumable;
            }
        }

        public bool CreateTypeConsumable(TypeConsumable instance)
        {
            if (instance.ID == 0)
            {
                Db.TypeConsumable.InsertOnSubmit(instance);
                Db.TypeConsumable.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool UpdateTypeConsumable(TypeConsumable instance)
        {
            TypeConsumable cache = Db.TypeConsumable.Where(p => p.ID == instance.ID).FirstOrDefault();
            if (cache != null)
            {
                //TODO : Update fields for Component
                Db.TypeConsumable.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool RemoveTypeConsumable(TypeConsumable instance)
        {
            try
            {
                Db.TypeConsumable.DeleteOnSubmit(instance);
                Db.TypeConsumable.Context.SubmitChanges();
                return true;
            }
            catch (Exception)
            {

                return false;
            }

        }
        #endregion
        

        #region Tag

        public IQueryable<Tag> Tags
        {
            get
            {
                return Db.Tag;
            }
        }

        public bool CreateTag(Tag instance)
        {
            if (instance.ID == 0)
            {
                Db.Tag.InsertOnSubmit(instance); 
                Db.Tag.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool UpdateTag(Tag instance)
        {
            Tag cache = Db.Tag.Where(p => p.ID == instance.ID).FirstOrDefault();
            if (cache != null)
            {
                //TODO : Update fields for Tag
                Db.Tag.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool RemoveTag(Tag instance)
        {
            try
            {
                Db.Tag.DeleteOnSubmit(instance);
                Db.Tag.Context.SubmitChanges();
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }
    #endregion


        # region SizePaper
                [Display(Name = "Типоразмер бумаги")]
                public IQueryable<SizePaper> SizePapers
                {
                    get { return Db.SizePaper; }
                }


                public bool CreateSizePaper(SizePaper instance)
                {

                    Db.SizePaper.InsertOnSubmit(instance);
                    Db.SizePaper.Context.SubmitChanges();
                    return true;
                }

                public bool UpdateSizePaper(SizePaper instance)
                {

                    Db.SizePaper.Context.SubmitChanges();
                    return true;
                }

                public bool RemoveSizePaper(SizePaper instance)
                {
                    try
                    {
                        Db.SizePaper.DeleteOnSubmit(instance);
                        Db.SizePaper.Context.SubmitChanges();
                        return true;
                    }
                    catch (Exception)
                    {

                        return false;
                    }

                }
        #endregion

        # region FieldSetting
                [Display(Name = "Поле настройки")]
                public IQueryable<Settings> FieldSettings
                {
                    get
                    {
                        return Db.Settings;
                    } 
                }

                public bool CreateFieldSetting(Settings instance)
                {
                    Db.Settings.InsertOnSubmit(instance);
                    Db.Settings.Context.SubmitChanges();
                    return true;
                }



                public bool UpdateFieldSetting(Settings instance)
                {
                    Db.Settings.Context.SubmitChanges();
                    return true;
                }

                public bool RemoveFieldSetting(Settings instance)
                {
                    Db.Settings.DeleteOnSubmit(instance);
                    Db.Settings.Context.SubmitChanges();
                    return true;
                }
         #endregion

        #region PaperType

                [Display(Name= "Тип бумаги")]
                public IQueryable<PaperType> PaperTypes
                {
                    get
                    {
                        return Db.PaperType;
                    }
                }

                public IQueryable<PaperType> PrinterPaperTypes  
                { 
                    get
                    {
                        //return Db.PaperType.Where(p => p.DeviceTypeID == 2);
                        return null;
                    }
                }
                public IQueryable<PaperType> PlotterPaperTypes 
                {
                    get
                    {
                        //     return Db.PaperType.Where(p => p.DeviceTypeID == 1);
                        return null;
                    }
                }


                public bool CreatePapertype(PaperType instance)
                {

                    Db.PaperType.InsertOnSubmit(instance);
                    Db.PaperType.Context.SubmitChanges();
                    return true;
                }

                public bool UpdatePapertype(PaperType instance)
                {
            
                    Db.PaperType.Context.SubmitChanges();
                    return true;
                }

        public bool RemovePapertype(PaperType instance)
        {
            try
            {
                Db.PaperType.DeleteOnSubmit(instance);
                Db.PaperType.Context.SubmitChanges();
                return true;
            }
            catch (Exception)
            {

                return false;

            }
        }

        #endregion


        #region PrintKind definition
                public IQueryable<PrintKind> PrintKinds
                {
                    get
                    {
                        return Db.PrintKind;
                    }
                }

                public bool CreatePrintKind(PrintKind instance)
                {
                    Db.PrintKind.InsertOnSubmit(instance);
                    Db.PrintKind.Context.SubmitChanges();
                    return true;
                }

                public bool UpdatePrintKind(PrintKind instance)
                {
                    Db.PrintKind.Context.SubmitChanges();
                    return true;
                }

                public bool RemovePrintKind(PrintKind instance)
                {
                    try
                    {
                        Db.PrintKind.DeleteOnSubmit(instance);
                        Db.PrintKind.Context.SubmitChanges();
                        return true;
                    }
                    catch (Exception)
                    {

                        return false;
                    }

                }
        #endregion


        #region Depertment

                public IQueryable<Department> Departments
                {
                    get
                    {
                        return Db.Department;
                    }
                }

                public bool CreateDepartment(Department instance)
                {
                    if (instance.ID == 0)
                    {
                        Db.Department.InsertOnSubmit(instance);
                        Db.Department.Context.SubmitChanges();
                        return true;
                    }

                    return false;
                }

                public bool UpdateDepartment(Department instance)
                {
                    Department cache = Db.Department.Where(p => p.ID == instance.ID).FirstOrDefault();
                    if (cache != null)
                    {
                        //TODO : Update fields for Department
                        Db.Department.Context.SubmitChanges();
                        return true;
                    }

                    return false;
                }

                public bool RemoveDepartment(Department instance)
                {
                    try
                    {
                        Db.Department.DeleteOnSubmit(instance);
                        Db.Department.Context.SubmitChanges();
                        return true;
                    }
                    catch (Exception)
                    {

                        return false;
                    }
                        

                }
        #endregion

                #region Manufacture


                public IQueryable<Manufacturer> Manufacturers
                {
                    get
                    {
                        return Db.Manufacturer;
                    }
                }

                public bool CreateManufacturer(Manufacturer instance)
                {
                    if (instance.ID == 0)
                    {
                        Db.Manufacturer.InsertOnSubmit(instance);
                        Db.Manufacturer.Context.SubmitChanges();
                        return true;
                    }

                    return false;
                }
        //автокомплит вывод списка
        public List<string> SearchManufacturer(string term)
        {
            return Db.Manufacturer.Where(m => m.Name.ToLower().Contains(term.ToLower())).Select(m => m.Name).ToList();
            
        }


        public List<string> SearchConsumble(string term,List<Consumable> consumables )
        {

            if (consumables != null) 
            {
                var z = Db.Consumable.Where(x => x.Name.ToLower().Contains(term.ToLower())).ToEnumerable();//..ToList();
                return z.Where(x => consumables.All(y => y.ID != x.ID)).Select(m => m.Name).ToList();
            }
            else
            {
                 return Db.Consumable.Where(m => m.Name.ToLower().Contains(term.ToLower())).Select(m => m.Name).ToList();

            }

        }                                              
        //проверка
        public int? CheckManufacturer(string name)
        {
            return Db.Manufacturer.Where(m => m.Name.ToLower() == name.ToLower()).Select(m => m.ID).SingleOrDefault();
        }

                public bool UpdateManufacturer(Manufacturer instance)
                {
                    Manufacturer cache = Db.Manufacturer.Where(p => p.ID == instance.ID).FirstOrDefault();
                    if (cache != null)
                    {
                        //TODO : Update fields for Manufacturer
                        Db.Manufacturer.Context.SubmitChanges();
                        return true;
                    }

                    return false;
                }

                public bool RemoveManufacturer(Manufacturer instance)
                {
                    try
                    {
                        Db.Manufacturer.DeleteOnSubmit(instance);
                        Db.Manufacturer.Context.SubmitChanges();
                        return true;
                    }
                    catch (Exception)
                    {

                        return false;
                    }

                }
        
#endregion
                #region Componenet
                [Display(Name = "Комплектующий")]
                public IQueryable<Component> Components
                {
                    get
                    {
                        return Db.Component;
                    }
                }

                public bool CreateComponent(Component instance)
                {
                    if (instance.ID == 0)
                    {
                        Db.Component.InsertOnSubmit(instance);
                        Db.Component.Context.SubmitChanges();
                        return true;
                    }

                    return false;
                }

                public bool UpdateComponent(Component instance)
                {
                    Component cache = Db.Component.Where(p => p.ID == instance.ID).FirstOrDefault();
                    if (cache != null)
                    {
                        //TODO : Update fields for Component
                        Db.Component.Context.SubmitChanges();
                        return true;
                    }

                    return false;
                }

                public bool RemoveComponent(Component instance)
                {
                   // Component instance = Db.Component.Where(p => p.ID == idComponent).FirstOrDefault();
                    if (instance != null)
                    {
                        Db.Component.DeleteOnSubmit(instance);
                        Db.Component.Context.SubmitChanges();
                        return true;
                    }

                    return false;
                }
                #endregion
        

        #region Application definition
                [Display(Name = "Приложения")]
                public IQueryable<Application> Applications
                {
                    get
                    {
                        return Db.Application;
                    }
                }

                public bool CreateApplication(Application instance)
                {
                    if (instance.ID == 0)
                    {
                        Db.Application.InsertOnSubmit(instance);
                        Db.Application.Context.SubmitChanges();
                        return true;
                    }

                    return false;
                }

                public bool UpdateApplication(Application instance)
                {
                    Application cache = Db.Application.Where(p => p.ID == instance.ID).FirstOrDefault();
                    if (cache != null)
                    {
                        //TODO : Update fields for Application
                        Db.Application.Context.SubmitChanges();
                        return true;
                    }

                    return false;
                }

                public bool RemoveApplication(Application instance)
                {
                    try
                    {
                        Db.Application.DeleteOnSubmit(instance);
                        Db.Application.Context.SubmitChanges();
                        return true;
                    }
                    catch (Exception)
                    {

                        return false;

                            
                        //error page
                    }
                       
                }
        #endregion

        #region CartridgeColor definition


                public IQueryable<CartridgeColor> CartridgeColors
                {
                    get
                    {
                        return Db.CartridgeColor;
                    }
                } //private set?
                public bool CreateCartridgeColor(CartridgeColor instance)
                {
                    Db.CartridgeColor.InsertOnSubmit(instance);
                    Db.CartridgeColor.Context.SubmitChanges();
                    return true;
                }

                public bool UpdateCartridgeColor(CartridgeColor instance)
                {
                    Db.CartridgeColor.Context.SubmitChanges();
                    return true;
                }

                public bool RemoveCartridgeColor(CartridgeColor instance)
                {
                    try
                    {
                        Db.CartridgeColor.DeleteOnSubmit(instance);
                        Db.CartridgeColor.Context.SubmitChanges();
                        return true;
                    }
                    catch (Exception)
                    {

                        return false;
                    }
                    
                }
        #endregion


        #region DeviceType definition

                [Display(Name = "Типы устройств")]
                public IQueryable<DeviceType> DeviceTypes
                {

                    get
                    {
                        return Db.DeviceType;            
                    }
                }



                public bool CreateDevicetype(DeviceType instance)
                {
                    Db.DeviceType.InsertOnSubmit(instance);
                    Db.DeviceType.Context.SubmitChanges();
                    return true;
                }
        
                public bool UpdateDevicetype(DeviceType instance)
                {
                    Db.DeviceType.Context.SubmitChanges();
                    return true;
                }

                public bool RemoveDevicetype(DeviceType instance)
                {
                    Db.DeviceType.DeleteOnSubmit(instance);
                    Db.DeviceType.Context.SubmitChanges();
                    return true;
                }
        #endregion

        
        #region Сотрудники-авторы
        [Display(Name = "Сотрудники-авторы")]
        public IQueryable<Employee> AuthorEmployees
        {
            get
            {
                return Db.Employee.Where(p => p.DepartmentID != 2);
            }
        }
        #endregion

        #region Employee definition
        [Display(Name = "Сотрудники")]
        public IQueryable<Employee> Employees
        {
            get
            {
                return Db.Employee;
            }
        }

        public bool CreateEmployee(Employee instance)
        {
            if (instance.ID == 0)
            {
                Db.Employee.InsertOnSubmit(instance);
                Db.Employee.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool UpdateEmployee(Employee instance)
        {
            Employee cache = Db.Employee.Where(p => p.ID == instance.ID).FirstOrDefault();
            if (cache != null)
            {
                //TODO : Update fields for Employee
                Db.Employee.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool RemoveEmployee(Employee instance)
        {

                Db.Employee.DeleteOnSubmit(instance);
                Db.Employee.Context.SubmitChanges();
                return true;

        }
#endregion
        

        #region сотрудники-пользователи
        [Display(Name = "Сотрудники-пользователи")]
        public IQueryable<Employee> UserEmployees
        {
            get
            {
                //==2 зачем ?
                //TODO
                return Db.Employee.Where(p => p.DepartmentID != 2);
            }
        }

        #endregion


        #region TagType

        public IQueryable<TagType> TagTypes
        {
            get
            {
                return Db.TagType;
            }
        }

        public bool CreateTagType(TagType instance)
        {
            if (instance.ID == 0)
            {
                Db.TagType.InsertOnSubmit(instance);
                Db.TagType.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool UpdateTagType(TagType instance)
        {
            TagType cache = Db.TagType.Where(p => p.ID == instance.ID).FirstOrDefault();
            if (cache != null)
            {
                //TODO : Update fields for TagType
                Db.TagType.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool RemoveTagType(TagType instance)
        {
            try
            {
                Db.TagType.DeleteOnSubmit(instance);
                Db.TagType.Context.SubmitChanges();
                return true;
            }
            catch (Exception)
            {
                
                return false;;
            }


            
        }

        #endregion
        


        //public IQueryable<Setup> Setup
        //{
        //    get
        //    {
        //        return Db.Setup;
        //    }
        //}

        //public bool SaveSettings(Setup settings)
        //{

        //    var s = Db.Setup.FirstOrDefault();
        //    if (s != null)
        //    {
        //        s.AccountName = settings.AccountName;
        //        s.MailServer = settings.MailServer;
        //        s.Pwd = settings.Pwd;
        //        s.Employee = Db.Employee.FirstOrDefault(p => p.TabNumber == "1190");
        //        s.Port = settings.Port;
        //    }
        //    else
        //    {
        //        settings.Employee = Db.Employee.FirstOrDefault(p => p.TabNumber == "1190");
        //        Db.Setup.InsertOnSubmit(settings);
        //    }

        //    Db.Setup.Context.SubmitChanges();
        //    return true;
        //}
    }
}