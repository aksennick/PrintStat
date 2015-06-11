using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.Globalization;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.IO.Compression;
using S22.Imap;
using System.Net.Mail;

namespace PrintStat.ParseModule
{
    public class SourceHPXml:Source
    {
        //public SourceHPXml()
        //{

        //}
        public override void parce(KeyValuePair<uint, MailMessage> message)
        {
            var context = new PrintStatDataDataContext();

            foreach (var attach in message.Value.Attachments)
            {
                ZipArchive z = new ZipArchive(attach.ContentStream);
                XmlDocument xml = new XmlDocument();
                foreach (var ent in z.Entries)
                {
                    xml.Load(ent.Open());

                    foreach (XmlElement x in xml.GetElementsByTagName("ACCOUNTING_INFO"))
                    //foreach (XmlElement x in xml.GetElementsByTagName("kmloginfo:print_job_log"))
                    {
                        var j = new Job();
                        //xml.GetElementsByTagName("Product_Name")[0].InnerText;//Принтер
                       

                        j.Name = GetAtribute(x, "JOB_NAME", "value");
                        try
                        {

                            j.ApplicationID = context.Application.FirstOrDefault(p => p.Name == GetAtribute(x, "APPLICATION_ID", "value")).ID;//a.ID;
                        }
                        catch (Exception ex)
                        { j.ApplicationID = context.Application.FirstOrDefault(p => p.Name == "Default").ID; }
                        j.Pages = Convert.ToInt32(GetAtribute(x, "PAGES", "value"));

                        j.Copies = Convert.ToInt32(GetAtribute(x, "COPIES", "value"));

                        j.StartTime = DateTime.ParseExact(GetAtribute(x, "PRINTING_TIMESTAMP", "value"), "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                        j.EndTime = DateTime.ParseExact(GetAtribute(x, "TIMESTAMP", "value"), "yyyyMMddHHmmss", CultureInfo.InvariantCulture);
                        //TIMESTAMP                        
                        j.Width_cm = (decimal)(Convert.ToInt32(GetAtribute(x, "WIDTH", "value")) * 0.0007);

                        j.Height_cm = (decimal)(Convert.ToInt32(GetAtribute(x, "LENGTH", "value")) * 0.0007);
                        try
                        {
                            j.UserTabNumber = context.Employee.First(p => p.ID == Convert.ToInt32(GetAtribute(x, "USER_NAME", "value"))).ID.ToString();
                        }
                        catch
                        {
                            j.UserTabNumber = "1369";
                        }
                        
             
                         j.DeviceID = context.Device.FirstOrDefault(p => p.SearchString == "hp").ID;

                        foreach (XmlElement el in x.GetElementsByTagName("CONSUME"))
                        {
                            try
                            {
                                var exforj = new ExpForJob();
                                exforj.CartridgeID = context.Cartridge.First(p => p.DeviceID == j.DeviceID && p.CartridgeColor.ShortName == el.Attributes[0].Value).ID; ;
                                exforj.JobID = j.ID;
                                exforj.Amount = Convert.ToInt32(el.Attributes[1].Value);
                                context.ExpForJob.InsertOnSubmit(exforj);
                            }
                            catch (Exception ex)
                            { 
                            }
                        }
                        context.Job.InsertOnSubmit(j);
                        context.SubmitChanges();


                    }
                }

                //XmlDocument xml = new XmlDocument();
                //xml.Load(attach.ContentStream);
            }
        }
    }
}