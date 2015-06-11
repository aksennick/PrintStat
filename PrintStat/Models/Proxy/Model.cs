using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using PrintStat;

namespace PrintStat
{
    public partial class Model : IBaseObject
    {
        public override string ToString()
        {
            return Name;
        }
    }
}