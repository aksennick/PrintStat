using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Ninject;
using System.Data.Linq;
using System.Web.Mvc;
using System.Web.UI.WebControls;


namespace PrintStat.Models
{
    public partial class SQLRepository: IRepository 
    {
        [Inject]
        public PrintStatDataDataContext Db {get; set;}

        public IQueryable<Device> PrintersAndPlotters
        {
            get
            {
                return Db.Device;
            }
        }

        public IQueryable<Device> Printers
        {
            get
            {
                return Db.Model.Join(Db.DeviceType.Where(dt => (dt.Name == "Принтер" || dt.Name == "принтер")),
                                    m => m.DeviceTypeID,
                                    dt => dt.ID,
                                    (m, dt) => m).Join(Db.Device,
                                                        d => d.ID,
                                                        dt => dt.ModelID,
                                                        (d, dt) => dt);

            }
        }

        public IQueryable<Device> Plotters
        {
            get
            {

                return Db.Model.Join(Db.DeviceType.Where(dt => (dt.Name=="Плоттер" || dt.Name=="плоттер")),
                                    m => m.DeviceTypeID,
                                    dt => dt.ID,
                                    (m, dt) => m).Join(Db.Device,
                                                        m => m.ID,
                                                        d => d.ModelID,
                                                        (m, d) => d);
            }
        }
        #region device
        public bool CreatePrinter(Device instance)
        {
             Db.Device.InsertOnSubmit(instance);
             Db.Device.Context.SubmitChanges();
             return true;
         }

        public bool UpdatePrinter(Device instance)
        {
            Device updating = Db.Device.FirstOrDefault(p => p.ID == instance.ID);
            if (updating != null)
            {
                updating.Name = instance.Name;
                Db.Device.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool RemovePrinter(Device instance)
        {
            Device deleting = Db.Device.FirstOrDefault(p=>p.ID == instance.ID);
            if (deleting != null)
            {
                Db.Device.DeleteOnSubmit(deleting);
                Db.Device.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public List<ModelConsumable> GetModelConsumables(int idDevice)
        {
           IQueryable<ModelConsumable> ids= Db.Device.Where(d => d.ID == idDevice).Join(
                Db.Model,
                d => d.ModelID,
                m => m.ID,
                (d, m) => m).Join(
                    Db.ModelConsumable,
                    m => m.ID,
                    mc => mc.ModelID,
                    (m, mc) => mc);
            return ids.ToList();

        }


        public IQueryable<DeviceConsumable> DeviceConsumables
        {
            get
            {
                return Db.DeviceConsumable;
            }
        }
        public class devCons
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public DateTime? DateInstalled { get; set; }
            public DateTime? DateEnd { get; set; }
            public string Type { get; set; }
            public bool? Uses { get; set; }
            public int Count { get; set; }
            public int CountSnmp { get; set; }
        }

        public int GetCountSnmp(int? idDevice,DateTime? start,DateTime? end)
        {

            //TODO сейчас идет только подсчет листов неважно какого формата
            //TODO привести к одному эквиваленту
            if (Db.Device.First(p => p.ID == idDevice).StatisticsSupported != true && end!=null)
            {
                
                return Convert.ToInt32((from d in Db.Device.Where(p => p.ID == idDevice)
                    join snmp in Db.SNMP.Where(l => l.Date >= start && l.Date <= end) on d.ID equals snmp.DeviceID
                    select snmp.Value).Sum());


            }
            else return 0;
        }

        public int GetCountPaper(int? idDevice, DateTime? start, DateTime? end)
        {

          //  TODO сделать подсчет как в страницах так и в погонных метрах

          // return (int)
            var i =
                (from d in Db.Device.Where(p => p.ID == idDevice)
                    join jb in Db.Job on d.ID equals jb.DeviceID
                    select new
                    {
                        jb.Copies,
                        jb.Pages,
                        RealHeight_m = jb.RealHeight_cm*10,
                        RealWidth_m = jb.RealWidth_cm*10

                    });
            if (Db.Device.First(p=>p.ID==idDevice).Model.DeviceType.Name=="Принтер")
            {
                return Convert.ToInt32(i.Select(p => p.Copies*p.Pages).Sum());
            }
            else
            {
                return Convert.ToInt32(i.Select(p => p.RealHeight_m).Sum());
            }
        }

        public List<devCons> GetDevConses(int? idDevice)
        {
            var devCon =
                from d in Db.Device.Where(p => p.ID == idDevice)
                join dc in Db.DeviceConsumable on d.ID equals dc.DeviceID
                join mc in Db.ModelConsumable on dc.ModelConsumableID equals mc.ID
                join c in Db.Consumable on mc.ConsumableID equals c.ID
                join tc in Db.TypeConsumable on c.TypeConsumableID equals tc.ID

                select new 
                {
                    deviceId=d.ID,
                    dc.Uses,
                    dc.ID,
                    c.Name,
                    dc.DateInstalled,
                    dc.DateEnd,
                    Type=tc.Name

                };
            List<devCons> listdc = new List<devCons>();
            foreach (var item in devCon)
            {
                var buf = new devCons();
                buf.ID = item.ID;
                buf.Name = item.Name;
                buf.DateInstalled = item.DateInstalled;
                buf.DateEnd = item.DateEnd;
                buf.Type = item.Type;
                buf.Uses = item.Uses;
                buf.Count = GetCountPaper(item.deviceId, item.DateInstalled, item.DateEnd);
                buf.CountSnmp = GetCountSnmp(item.deviceId, item.DateInstalled, item.DateEnd);

                listdc.Add(buf);
            }
            return listdc ;
        }
        public bool CreateDeviceConsumable(int idDevice,int idModelConsumable,DateTime time)
        {
            try
            {
                if (time == DateTime.MinValue) time = DateTime.Now;
                var t = DateTime.Now;
                var t1 = DateTime.Today;
                var instance = new DeviceConsumable()
                {
                    DeviceID = idDevice,
                    ModelConsumableID = idModelConsumable,
                    DateInstalled = time,
                    Uses = true
                };
                Db.DeviceConsumable.InsertOnSubmit(instance);
                Db.DeviceConsumable.Context.SubmitChanges();
                return true;
                
            }
            catch (Exception)
            {
                
                return false;
            }
        }

        //1 option
        public bool SetDateEnd(int devConId, DateTime dateEnd)
        {
            DeviceConsumable devcon = Db.DeviceConsumable.First(p => p.ID == devConId);
            if (dateEnd == DateTime.MinValue)
            {
                devcon.DateEnd = DateTime.Now;
            }
            else
            {
                devcon.DateEnd = dateEnd;
            }
            Db.DeviceConsumable.Context.SubmitChanges();

            return true;
        }
        //2 option
        public bool SetUseOffAndAddNewCons(int idDevCon)
        {
            //прекращаем использование
            DeviceConsumable temp = Db.DeviceConsumable.First(p => p.ID == idDevCon);
            temp.Uses = false;
            Db.DeviceConsumable.Context.SubmitChanges();
            //Заменяем на новую такую же?
            //TODO уточнить на такую же или на другую?
            CreateDeviceConsumable(temp.DeviceID, temp.ModelConsumableID, DateTime.MinValue);

            return true;
        }

        public bool RemoveDeviceConsumable(int idDevice)
        {
            IEnumerable<DeviceConsumable> instance = Db.DeviceConsumable.Where(p => p.DeviceID == idDevice);
            if (instance != null)
            {
                Db.DeviceConsumable.DeleteAllOnSubmit(instance);
                Db.DeviceConsumable.Context.SubmitChanges();
                return true;
            }
            return false;
        }
#endregion

        public List<Model> SearchModel(string term)
        {
            if (term != "")
            {
                return
                    Db.Model.Where(m => m.Name.ToLower().Contains(term.ToLower())).ToList();

            }
            return Db.Model.ToList();
        }
        public List<Model> SerchModels(string name)
        {
            if (name!=null)
            {
                return  Db.Model.Where(m => m.Manufacturer.Name.ToLower() == name.ToLower()).ToList();
            }
            return null;
        }

        #region ModelRelation



        public bool RemoveModelConsumable(IQueryable<ModelConsumable> instance)
        {
            try
            {
                if (instance != null)
                {
                    Db.ModelConsumable.DeleteAllOnSubmit(instance);
                    Db.ModelConsumable.Context.SubmitChanges();
                    return true;
                }
                return false;
            }
            catch (Exception)
            {

                return false;
            }
        }

        public int[] CreateModelComsumable(int[] comIDs, int modId)//поменять местами
        {
            try
            {
                var modelConsList = (comIDs.Select(id =>
                    new ModelConsumable()
                    {
                        ConsumableID = id,
                        ModelID = modId
                    })).ToList();
               
                Db.ModelConsumable.InsertAllOnSubmit(modelConsList);
                Db.ModelConsumable.Context.SubmitChanges();
                return modelConsList.Select(p=>p.ID).ToArray();
            }
            catch (Exception)
            {
                
                return null;
            }

        }
        //todo удаление и добавление 
        //todo если удалять то спрашивать затем удалить все из dc потом уже из mc
        //
        //public bool UpdateModelConsumbles(int modelId, int[] comIDs)
        //{
            
        //    var mc = from p in Db.ModelConsumable.Where(p => p.ModelID == modelId)
        //        select new {idCon = p.ID, selected=false};

        //    foreach (var item in comIDs)
        //    {
        //        //if(mc.C)
        //    }
        //    return true;
        //}
        //public IQueryable<Consumable> GetNoInstallCons(int idDevice)
        //{
        //    Db.Device.Where(p=>p.ID==idDevice)
        //        .Join(Db.Model,d=>d.ModelID,m=>m.ID,(d,m)=>m)
        //        .Join(Db.ModelConsumable,m=>m.ID,mc=>mc.ModelID,(m,mc)=>mc)
        //}

        public IQueryable<ModelTag> ModelTags
        {
            get
            {
                return Db.ModelTag;
            }
        }

        public IQueryable<SupportSize> ModelSizePapers
        {
            get
            {
                return Db.SupportSize;
            }
        }

        public IQueryable<ModelPaperType> ModelPaperTypes
        {
            get
            {
                return Db.ModelPaperType;
            }
        }


        public bool CreateModelTag(int[] ChosenTagIds, int idModel)
        {
            if (ChosenTagIds != null)
            {
                Db.ModelTag.InsertAllOnSubmit(
                    ChosenTagIds.Select(id =>
                        new ModelTag
                        {
                            TagID = id,
                            ModelID = idModel
                        })
                    );
                Db.ModelTag.Context.SubmitChanges();
                return true;
            }
            return false;
        }

        public bool CreateModelSizePaper(int[] ChosenSizePaperIds, int idModel)
        {
            if (ChosenSizePaperIds != null)
            {
                Db.SupportSize.InsertAllOnSubmit(
                    ChosenSizePaperIds.Select(id =>
                        new SupportSize
                        {
                            SizePaperID = id,
                            ModelID = idModel
                        })
                    );
                Db.SupportSize.Context.SubmitChanges();
                return true;
            }
            return false;
        }

        public bool CreateModelPaperType(int[] ChosenPaperTypeIds, int idModel)
        {
            if (ChosenPaperTypeIds != null)
            {
                Db.ModelPaperType.InsertAllOnSubmit(
                    ChosenPaperTypeIds.Select(id =>
                        new ModelPaperType
                        {
                            PaperTypeID = id,
                            ModelID = idModel
                        })
                    );
                Db.ModelPaperType.Context.SubmitChanges();
                return true;
            }
            return false;
        }

        public bool RemoveModelTag(IQueryable<ModelTag> instance)
        {

            if (instance != null)
            {
                Db.ModelTag.DeleteAllOnSubmit(instance);
                Db.ModelTag.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool RemoveModelSizePaper(IQueryable<SupportSize> instance)
        {

            if (instance != null)
            {
                Db.SupportSize.DeleteAllOnSubmit(instance);
                Db.ModelTag.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool RemoveModelPaperType(IQueryable<ModelPaperType> instance)
        {

            if (instance != null)
            {
                Db.ModelPaperType.DeleteAllOnSubmit(instance);
                Db.ModelTag.Context.SubmitChanges();
                return true;
            }
            return false;
        }
#endregion

        #region Model

        public IQueryable<Model> Models
        {
            get
            {
                return Db.Model;
            }
        }

        public void ClearContextOfDeletes()
        {
            try
            {
                  Db.ModelConsumable.Context.GetChangeSet().Deletes.Clear();
            }
            catch (Exception)
            {
              return;
            }
          
        }

        public void SubmitContexOfDeletes()
        {
            try
            {
                Db.ModelConsumable.Context.SubmitChanges();
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        public bool CreateModel(Model instance)
        {
            if (instance.ID == 0)
            {
                Db.Model.InsertOnSubmit(instance);
                Db.Model.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool UpdateModel(Model instance)
        {
            Model cache = Db.Model.Where(p => p.ID == instance.ID).FirstOrDefault();
            if (cache != null)
            {
                //TODO : Update fields for Model
                Db.Model.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool RemoveModel(Model instance)
        {
            
            if (instance != null)
            {
                Db.Model.DeleteOnSubmit(instance);
                Db.Model.Context.SubmitChanges();
                return true;
            }

            return false;
        }
        
        #endregion
        

        public int? CheckCartridge(string temp)
        {
            try
            {
                return Db.TypeConsumable.Where(m => m.Name.ToLower().Contains(temp.ToLower())).Select(m => m.ID).FirstOrDefault();
            }
            catch (Exception)
            {
                
               return 0; 
            }
        }
        #region Consumable
        public IQueryable<Consumable> Consumables  
        {
            get
            {
                return Db.Consumable;
            }
        }

        public bool CreateConsumable(Consumable instance)    
        {
            if (instance.ID == 0)
            {
                Db.Consumable.InsertOnSubmit(instance);
                Db.Consumable.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public bool UpdateConsumable(Consumable instance)    
        {
            if (instance.ID == 0)
            {
                Db.Consumable.InsertOnSubmit(instance);
                Db.Consumable.Context.SubmitChanges();
                return true;
            }

            return false;
        }

        public int CountUsesConsumble(int? idConsumble)
        {
           return  Db.Consumable.Where(p => p.ID == idConsumble)
                .Join(Db.ModelConsumable,
                    c => c.ID,
                    mc => mc.ConsumableID,
                    (c, mc) => mc)
                .Join(Db.DeviceConsumable,
                    mc => mc.ID,
                    dc => dc.ModelConsumableID,
                    (mc, dc) => dc)
                .Count();
        }


        public bool RemoveConsumable(Consumable instance)
        {
            try
            {

                    Db.Consumable.DeleteOnSubmit(instance);
                    Db.Consumable.Context.SubmitChanges();
                    return true;

                
            }
            catch (Exception)
            {

                return false;
            }
            
        }



#endregion

        #region ModelConsumable


        public IQueryable<ModelConsumable> ModelConsumables
        {
            get
            {
                return Db.ModelConsumable;
            }
        }

        

        #endregion

    }
}