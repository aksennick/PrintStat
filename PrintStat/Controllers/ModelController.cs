using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using PrintStat.Models.ViewModels;

namespace PrintStat.Controllers
{
        public class ModelController : BaseController
        {
            //
            // GET: /CartridgeColor/
            List<Consumable> _Consumables = new List<Consumable>();

            List<Consumable> Consumables
            {
                get { return _Consumables; }
                set { _Consumables = value; }
            }

            [HttpGet]
            public PartialViewResult DelModelCons(int? idMod,int idCons)
            {

                        var temp = Session["Cons"] as List<Consumable>;
                if (temp==null)
                {
                    temp = Repository.ModelConsumables.Where(p => p.ModelID == idMod).Join(
                        Repository.Consumables,
                        p => p.ConsumableID,
                        x => x.ID,
                        (p, x) => x
                        ).ToList();
                }
                        if (temp!=null)
                        {
                            temp.RemoveAll(x=>x.ID==idCons);
                            Session["Cons"] = temp;
                        }
                        return PartialView("partialModelConsumable", temp);

                return null;
            }

            [HttpGet]
            public ActionResult AddConsToModel(string name)
                {
                    var temp = new List<Consumable>();
                    if(Session["Cons"]!=null)
                    temp = Session["Cons"] as List<Consumable>;

                var Cons = Repository.Consumables.FirstOrDefault(p => p.Name == name);
                if (Cons != null)
                {
                        temp.Add(Cons);
                        Session["Cons"] = temp;
                        return PartialView("partialModelConsumable", temp);
                }
                return PartialView("partialModelConsumable", temp);
            }
            public ActionResult Index()
            {
                var model = Repository.Models.ToList();
                return View(model);
            }


            private void InitViewBag()
            {
                ViewBag.Tags = Repository.Tags;
                ViewBag.DeviceTypes = Repository.DeviceTypes;
                ViewBag.PrintKinds = Repository.PrintKinds;
                ViewBag.Manufacturers = Repository.Manufacturers;
                ViewBag.Consumables = Repository.Consumables;
                
                ViewBag.PaperTypes = Repository.PaperTypes;
                ViewBag.SizePapers = Repository.SizePapers;


            }
            [HttpGet]
            public ActionResult CreateModel()
            {
                InitViewBag();
                var newModelView =  new ModelView();

     
                return View(newModelView);
            }


            [HttpPost]
            public ActionResult CreateModel(ModelView modelView)
            {
                var anyModel = Repository.Models.Any(p => String.Compare(p.Name, modelView.Name) == 0);
                if (anyModel)
                {
                    ModelState.AddModelError("Name", "Модель с таким наименованием уже существует");
                }

                if (ModelState.IsValid)
                {
                    var model = (Model)ModelMapper.Map(modelView, typeof(ModelView), typeof(Model));
                    Repository.CreateModel(model);
                    var modelId = model.ID;
                    var tempCons = Session["Cons"] as List<Consumable>;
                    if (tempCons != null)
                        Repository.CreateModelComsumable(tempCons.Select(x => x.ID).ToArray(), modelId);
                     Session.Abandon();
                    Repository.CreateModelTag(modelView.ChosenTagIds, modelId);
                    Repository.CreateModelPaperType(modelView.ChosenPaperTypeIds, modelId);
                    Repository.CreateModelSizePaper(modelView.ChosenSizePaperIds, modelId);
                    return RedirectToAction("Index");
                }
                InitViewBag();
                return View(modelView);
            }

            public ActionResult PartialCreateModel()
            {
                InitViewBag();
                return PartialView("PartialCreateModel");
            }



            public void SubmitContextDelete(int idMod)
            {
                var temp = Session["Cons"] as List<Consumable>;
                var usesModCons = Repository.ModelConsumables.Where(p => p.ModelID == idMod).ToList(); //здесь все что были

                try
                {
                    foreach (var item in usesModCons)
                    {
                        if (!temp.Any(p => p.ID == item.ConsumableID))
                        {
                            var item1 = item;
                            Repository.RemoveModelConsumable(
                                Repository.ModelConsumables.Where(
                                    p => p.ConsumableID == item1.ConsumableID && p.ModelID == idMod));
                        }
                    }
                }
                catch
                {
                    // ignored
                }
            }

            [HttpGet]
            public ActionResult EditModel(int? id)
            {
                InitViewBag();
                ViewBag.ID = id;
                var model = Repository.Models.FirstOrDefault(p => p.ID == id);
             
                if (model != null)
                {
                    var modelView = (ModelView)ModelMapper.Map(model, typeof(Model), typeof(ModelView));

                    var choosModelCons = Repository.ModelConsumables.Where(p => p.ModelID == id);
                    foreach (var item in choosModelCons)
                    {
                        modelView.Consumables.Add(Repository.Consumables.First(c=>c.ID == item.ConsumableID));
                    }
                    Session["Cons"] = modelView.Consumables;
                    var choosModelTag = Repository.ModelTags.Where(p => p.ModelID == id);
                    foreach (var item in choosModelTag)
                    {
                        modelView.Tags.Add(Repository.Tags.First(t=>t.ID==item.TagID));
                    }

                    var choosModelSizePaper = Repository.ModelSizePapers.Where(p => p.ModelID == id);
                    foreach (var item in choosModelSizePaper)
                    {
                        modelView.SizePapers.Add(Repository.SizePapers.First(t => t.ID == item.SizePaperID));
                    }

                    var choosModelPaperType = Repository.ModelPaperTypes.Where(p => p.ModelID == id);
                    foreach (var item in choosModelPaperType)
                    {
                        modelView.PaperTypes.Add(Repository.PaperTypes.First(t => t.ID == item.PaperTypeID));
                    }

                    return View(modelView);
                }
                return RedirectToAction("Index");
            }

            [HttpPost]
            public ActionResult EditModel(ModelView modelView)
            {
                if (ModelState.IsValid)
                {
                    var model = Repository.Models.FirstOrDefault(p => p.ID == modelView.ID);
                    ModelMapper.Map(modelView, model, typeof(ModelView), typeof(Model));
                    Repository.UpdateModel(model);
                    var modelId = modelView.ID;
                   
                    var modelTag = Repository.ModelTags.Where(mt => mt.ModelID == modelId);
                    Repository.RemoveModelTag(modelTag);
                    Repository.CreateModelTag(modelView.ChosenTagIds, modelId);

                    var modelSizePaper = Repository.ModelSizePapers.Where(mt => mt.ModelID == modelId);
                    Repository.RemoveModelSizePaper(modelSizePaper);
                    Repository.CreateModelSizePaper(modelView.ChosenSizePaperIds, modelId);

                    var modelPaperType = Repository.ModelPaperTypes.Where(mt => mt.ModelID == modelId);
                    Repository.RemoveModelPaperType(modelPaperType);
                    Repository.CreateModelPaperType(modelView.ChosenPaperTypeIds, modelId);


                    var cons = Session["Cons"] as List<Consumable>;
                    Session.Abandon();
                    var modelCons = Repository.ModelConsumables.Where(mc => mc.ModelID == modelId);
                    int[] idModCons =Repository.CreateModelComsumable((from item in cons where !modelCons.Any(p => p.ConsumableID == item.ID) select item.ID).ToArray(), modelId);
                    int[] idDeviceOfModel =
                        Repository.PrintersAndPlotters.Where(p => p.ModelID == modelView.ID).Select(p => p.ID).ToArray();
                    foreach (var itemModCon in idModCons)
                    {
                        foreach (var itemDevice in idDeviceOfModel)
                        {
                            Repository.CreateDeviceConsumable(itemDevice,itemModCon,DateTime.MinValue);
                        }
                    }
                    
                    return RedirectToAction("Index");
                }

                return View(modelView);
            }


            [HttpGet]
            public ActionResult DeleteModel(int? id)
            {
                var model = Repository.Models.FirstOrDefault(p => p.ID == id);
                if (model != null)
                {
                    try
                    {
                        var modelConsumable = Repository.ModelConsumables.Where(mt => mt.ModelID == id);
                        Repository.RemoveModelConsumable(modelConsumable);
                        //todo удалить mc если есть 
                        if (Repository.RemoveModel(model))
                        {
                            var modelPaperType = Repository.ModelPaperTypes.Where(mt => mt.ModelID == id);
                            Repository.RemoveModelPaperType(modelPaperType);

                            var modelPaperSizePaper = Repository.ModelSizePapers.Where(mt => mt.ModelID == id);
                            Repository.RemoveModelSizePaper(modelPaperSizePaper);

                            var modelPaperTag = Repository.ModelTags.Where(mt => mt.ModelID == id);
                            Repository.RemoveModelTag(modelPaperTag);


                        }
                    }
                    catch (Exception)
                    {
                        ModelState.AddModelError("","Невозможно удалить модель "+model.Manufacturer.Name+" "+ model.Name+", так как она используется");
                        var _models = Repository.Models.ToList();
                        return View("Index",_models);
                    }
                
                }
                return RedirectToAction("Index");
            }

            struct _model
            {
                public int ID;
                public string Name;
            }
            struct __model
            {
                public string i;
                public string value;
                public string text;
            }
            public string GetModelForManufacturer(string id)
            {
                
                _model tempModel = new _model();
                var _models = new List<_model>();
                if (id!=null)
                {
                    var models = new List<Model>();
                    models = Repository.SerchModels(id);

                    if (models.Count!=0)
                    {
                        foreach (var m in models)
                        {
                            tempModel.ID = m.ID;
                            tempModel.Name = m.Name;
                            _models.Add(tempModel);
                        }
                    }
                    else
                    {
                        tempModel.ID = 0;
                        tempModel.Name = "Моделей нет";
                        _models.Add(tempModel);
                    }
                    
                    
                }
                    else
                    {
                        tempModel.ID = 0;
                        tempModel.Name = "Моделей нет";
                        _models.Add(tempModel);
                    }

                    return new JavaScriptSerializer().Serialize(_models);
            }

           
            //public JsonResult GetModels(string search)
            //{
            //    __model tempModel = new __model();
            //    var _models = new List<__model>();
            //    var models = new List<Model>();
            //    models = Repository.Models.ToList();
            //    if (models.Count != 0)
            //    {
            //        int i = 1;
            //        foreach (var m in models)
            //        {
            //            tempModel.i = i.ToString();
            //            tempModel.value = m.ID.ToString();
            //            tempModel.text = m.Name;
            //            _models.Add(tempModel);
            //            i++;
            //        }
            //    }
            //    return Json(_models, JsonRequestBehavior.AllowGet);
                 
            //}


            public JsonResult AutocompleteManufacturer(string term)
            {

                return Json(Repository.SearchManufacturer(term), JsonRequestBehavior.AllowGet);

            }
  
            [HttpPost]
            public bool AddModel(ModelView objModelView)
            {
                var anyModel = Repository.Models.Any(p => String.Compare(p.Name,objModelView.Name)==0);
                if (!anyModel)
                {
                    var _model = new Model()
                    {
                        Name = objModelView.Name,
                        ManufacturerID = objModelView.ManufacturerID,
                        DeviceTypeID = objModelView.DeviceTypeID,
                        PrintKindID = objModelView.PrintKindID
                    };
                    Repository.CreateModel(_model);
                    var modelId = _model.ID;

                    Repository.CreateModelTag(objModelView.ChosenTagIds, modelId);
                    Repository.CreateModelPaperType(objModelView.ChosenPaperTypeIds, modelId);
                    Repository.CreateModelSizePaper(objModelView.ChosenSizePaperIds, modelId);
                    return true;
                }

                return false;
            }

            //public int? CreateModel(string modelView)
            //{
            //    var anyModel = Repository.Models.Any(p => String.Compare(p.Name, modelView.Name) == 0);
            //    if (anyModel)
            //    {
            //        ModelState.AddModelError("Name", "Модель с таким наименованием уже существует");
            //    }

            //    if (ModelState.IsValid)
            //    {

            //        var model = (Model)ModelMapper.Map(modelView, typeof(ModelView), typeof(Model));
            //        Repository.CreateModel(model);
            //        var modelId = model.ID;
            //        Repository.CreateModelTag(modelView.ChosenTagIds, modelId);
            //        Repository.CreateModelPaperType(modelView.ChosenPaperTypeIds, modelId);
            //        Repository.CreateModelSizePaper(modelView.ChosenSizePaperIds, modelId);
            //        return RedirectToAction("Index");
            //    }

            //    return View(modelView);
            //}
        }


        
    }

