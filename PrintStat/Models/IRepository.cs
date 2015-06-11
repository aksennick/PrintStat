using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using PrintStat.Models.ViewModels;

namespace PrintStat.Models
{
    public interface IRepository
    {

        #region Model

        void ClearContextOfDeletes();
        void SubmitContexOfDeletes();
        IQueryable<Model> Models { get; }

        bool CreateModel(Model instance);

        bool UpdateModel(Model instance);

        bool RemoveModel(Model instance);

        List<Model> SerchModels(string name);
        #endregion 
        

        #region ModelRelation
        //комплектующие
        bool RemoveModelConsumable(IQueryable<ModelConsumable> instance);
        int[] CreateModelComsumable(int[] comIDs, int modelID);
        //bool UpdateModelConsumbles(int modelId,int[] comIDs);

        IQueryable<ModelTag> ModelTags { get; }
        IQueryable<SupportSize> ModelSizePapers { get; }
        IQueryable<ModelPaperType> ModelPaperTypes { get; }

        //bool CheckAsignTag(int modelId, int tagId);

        bool CreateModelTag(int[] ChosenTagIds, int idModel);
        bool CreateModelPaperType(int[] ChosenPaperTypeIds, int idModel);
        bool CreateModelSizePaper(int[] ChosenSizePaperIds, int idModel);

        bool RemoveModelTag(IQueryable<ModelTag> instance);

        bool RemoveModelSizePaper(IQueryable<SupportSize> instance);

        bool RemoveModelPaperType(IQueryable<ModelPaperType> instance);

        #endregion 
        

        #region Manufacturer

        IQueryable<Manufacturer> Manufacturers { get; }

        bool CreateManufacturer(Manufacturer instance);

        bool UpdateManufacturer(Manufacturer instance);

        bool RemoveManufacturer(Manufacturer instance);

        List<string> SearchManufacturer(string term);
        //List<string> SearchConsumble(string term);
        List<string> SearchConsumble(string term, List<Consumable> consumables);

        List<Model> SearchModel(string term);

        int? CheckManufacturer(string name);
        #endregion 
        

        #region Component

        IQueryable<Component> Components { get; }

        bool CreateComponent(Component instance);

        bool UpdateComponent(Component instance);

        bool RemoveComponent(Component instance);

        #endregion 

        #region TypeConsumable

        IQueryable<TypeConsumable> TypeConsumables { get; }

        bool CreateTypeConsumable(TypeConsumable instance);

        bool UpdateTypeConsumable(TypeConsumable instance);

        bool RemoveTypeConsumable(TypeConsumable instance);

        #endregion 
        


        #region Consumable


        int? CheckCartridge(string temp);
        IQueryable<Consumable> Consumables { get; }

        bool CreateConsumable(Consumable instance);

        bool UpdateConsumable(Consumable instance);

        bool RemoveConsumable(Consumable instance);
        int CountUsesConsumble(int? idConsumble);

        List<SQLRepository.devCons> GetDevConses(int? idDevice);

        bool SetDateEnd(int devConId, DateTime dateEnd);
        bool SetUseOffAndAddNewCons(int idDevCon);
       // int GetCountSnmp(int? idDevice, DateTime start, DateTime? end);
        #endregion 


        #region ModelConsumable

        IQueryable<ModelConsumable> ModelConsumables { get; }



        #endregion 
        

        


        IQueryable<Device> PrintersAndPlotters { get;  }
        IQueryable<Device> Printers { get; }
        IQueryable<Device> Plotters { get; }

        //Добавление 
        IQueryable<SizePaper> SizePapers { get; }
        IQueryable<CartridgeColor> CartridgeColors { get;}

        IQueryable<PaperType> PaperTypes { get;  }
        IQueryable<PaperType> PrinterPaperTypes { get;  }
        IQueryable<PaperType> PlotterPaperTypes { get; }
        
        IQueryable<DeviceType> DeviceTypes { get;  }

        IQueryable<PrintKind> PrintKinds { get;  }
        IQueryable<Job> Jobs { get;  }



        IQueryable<Employee> AuthorEmployees { get;  }
        IQueryable<Employee> UserEmployees { get; }

       // IQueryable<Setup> Setup { get; }
//Setting
        #region Profile

        IQueryable<Profile> Profiles { get; }

        bool CreateProfile(Profile instance);

        bool UpdateProfile(Profile instance);

        bool RemoveProfile(Profile instance);

        #endregion 

        #region Settings

        IQueryable<Settings> Settingses { get; }

        bool CreateSettings(Settings instance);

        bool UpdateSettings(Settings instance);

        bool RemoveSettings(Settings instance);

        #endregion 
        
        #region SettingsValue

        IQueryable<SettingValue> SettingValues { get; }

        bool CreateSettingValue(int[] setId, int profId);

        bool UpdateSettingValue(List<ProfileView.setval> sv);

        bool RemoveSettingValue(SettingValue instance);

        #endregion 
 /// <summary>
 /// ///
 /// </summary>
       
        #region FieldSettings

        IQueryable<Settings> FieldSettings { get; }

        bool CreateFieldSetting(Settings instance);

        bool UpdateFieldSetting(Settings instance);

        bool RemoveFieldSetting(Settings instance);

        #endregion 


        #region Employee

        IQueryable<Employee> Employees { get; }

        bool CreateEmployee(Employee instance);

        bool UpdateEmployee(Employee instance);

        bool RemoveEmployee(Employee instance);

        #endregion 
        

        #region Application

        IQueryable<Application> Applications { get; }

        bool CreateApplication(Application instance);

        bool UpdateApplication(Application instance);

        bool RemoveApplication(Application instance);

        #endregion 
        

        #region Department

        IQueryable<Department> Departments { get; }

        bool CreateDepartment(Department instance);

        bool UpdateDepartment(Department instance);

        bool RemoveDepartment(Department instance);

        #endregion 
        

        #region TagType

        IQueryable<TagType> TagTypes { get; }

        bool CreateTagType(TagType instance);

        bool UpdateTagType(TagType instance);

        bool RemoveTagType(TagType instance);

        #endregion 
        
        #region Device
        bool CreatePrinter(Device instance);
        bool UpdatePrinter(Device instance);
        bool RemovePrinter(Device instance);



        List<ModelConsumable> GetModelConsumables(int idDevice);


        #region DeviceConsumable

        IQueryable<DeviceConsumable> DeviceConsumables { get; }

        bool CreateDeviceConsumable(int idDevice, int idModelConsumable, DateTime time);

        bool RemoveDeviceConsumable(int idDevice);

        #endregion 
        
#endregion


        bool CreatePapertype(PaperType instance);
        bool UpdatePapertype(PaperType instance);

        bool RemovePapertype(PaperType instance);

        // SizePaper
        bool CreateSizePaper(SizePaper instance);
        bool UpdateSizePaper(SizePaper instance);
        bool RemoveSizePaper(SizePaper instance);
        //PrintKind
        bool CreatePrintKind(PrintKind instance);
        bool UpdatePrintKind(PrintKind instance);
        bool RemovePrintKind(PrintKind instance);

        //CartridgeColor
        bool CreateCartridgeColor(CartridgeColor instance);
        bool UpdateCartridgeColor(CartridgeColor instance);
        bool RemoveCartridgeColor(CartridgeColor instance);

        bool CreateDevicetype(DeviceType instance);
        bool UpdateDevicetype(DeviceType instance);

        bool RemoveDevicetype(DeviceType instance);

        bool CreateJob(Job instance);
        bool UpdateJob(Job instance);
        bool RemoveJob(Job instance);

       // bool SaveSettings(Setup settings);



        #region Tag

        IQueryable<Tag> Tags { get; }

        bool CreateTag(Tag instance);

        bool UpdateTag(Tag instance);

        bool RemoveTag(Tag instance);

        #endregion 
        

        

    }
}
