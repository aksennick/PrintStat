using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;


namespace PrintStat.Models
{
    public partial class SQLRepository: IRepository
    {

        [Display(Name = "Задания")]
        public IQueryable<Job> Jobs
        {
            get
            {
                return Db.Job.OrderBy(p=>p.StartTime);
            }
        }


        public bool CreateJob(Job instance)
        {

            Db.Job.InsertOnSubmit(instance);
            Db.Job.Context.SubmitChanges();
            return true;
        }

        public bool UpdateJob(Job instance)
        {

            Db.Job.Context.SubmitChanges();
            return true;
        }

        public bool RemoveJob(Job instance)
        {
            Db.Job.DeleteOnSubmit(instance);
            Db.Job.Context.SubmitChanges();
            return true;
        }
    }
}