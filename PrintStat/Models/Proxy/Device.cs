using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PrintStat
{

    public partial class Device: IBaseObject
    {

        public override string ToString()
        {
            return Name;
        }

        //public DeviceType DeviceType
        //{
        //    get
        //    {
        //        return DeviceType;//
        //    }
        //}
    }
}