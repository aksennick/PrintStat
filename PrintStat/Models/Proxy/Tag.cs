using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PrintStat
{
    public partial class Tag : IBaseObject
    {
        public override string ToString()
        {
            return Name;
        }
        //public string Tag1
        //{
        //    get { return Tag1; }
        //}
    }
}