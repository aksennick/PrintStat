using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DotNetOpenAuth.OpenId.Extensions.AttributeExchange;

namespace PrintStat.Models.ViewModels
{
    public class ModelConsumableView : BaseView
    {
        

        public List<Consumable> ConsItem { get; set; } 
    }
}
