﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace PrintStat
{
    public partial class PaperType: IBaseObject
    {

        public override string ToString()
        {
            return Name;
        }
    }
}