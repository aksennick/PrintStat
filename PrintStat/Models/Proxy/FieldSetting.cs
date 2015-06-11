using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PrintStat
{
    public partial class Settings : IBaseObject
    {
        public override string ToString()
        {
            return Name;
        }
    }
}
