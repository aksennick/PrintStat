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
    public class SourceFSXml:Source
    {
        //public SourceHPXml()
        //{

        //}
        public override void parce(KeyValuePair<uint, MailMessage> message)
        {

            var context = new PrintStatDataDataContext();

            foreach (var attach in message.Value.Attachments)
            {
                XmlDocument xml = new XmlDocument();

                xml.Load(attach.ContentStream);

                foreach (XmlElement x in xml.GetElementsByTagName("kmloginfo:print_job_log"))
                {
                    var j = new Job();
                    //xml.GetElementsByTagName("Product_Name")[0].InnerText;//Принтер
                    j.DeviceID = context.Device.FirstOrDefault(p => p.SearchString == "FS").ID;
                    j.Name = GetInnerText(x, "kmloginfo:job_name");
                   
                    j.Pages = Convert.ToInt32(GetInnerText(x,"kmloginfo:complete_pages"));

                    j.Copies = Convert.ToInt32(GetInnerText(x, "kmloginfo:complete_copies"));


                    j.StartTime = new DateTime(Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:start_time")[0], "kmloginfo:year")),
                                                    Convert.ToInt32( GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:start_time")[0], "kmloginfo:month")) ,
                                                    Convert.ToInt32(  GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:start_time")[0], "kmloginfo:day")),
                                                    Convert.ToInt32(  GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:start_time")[0], "kmloginfo:hour")),
                                                    Convert.ToInt32(  GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:start_time")[0], "kmloginfo:minute")),
                                                    Convert.ToInt32(  GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:start_time")[0], "kmloginfo:second")));
                    //j.StartTime = DateTime.ParseExact(dat,"yyMMddHHmmss", CultureInfo.InvariantCulture);

                    j.EndTime = new DateTime(Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:end_time")[0], "kmloginfo:year")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:end_time")[0], "kmloginfo:month")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:end_time")[0], "kmloginfo:day")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:end_time")[0], "kmloginfo:hour")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:end_time")[0], "kmloginfo:minute")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName("kmloginfo:end_time")[0], "kmloginfo:second")));
                    try
                    {
                        j.UserTabNumber = context.Employee.FirstOrDefault(p => p.ID == Convert.ToInt32(GetInnerText(x, "kmloginfo:user_name"))).ID.ToString();
                    }
                    catch
                    {
                        j.UserTabNumber = "1369";
                    }
                    j.ApplicationID = context.Application.First(p => p.Name == "Default").ID;
                    context.Job.InsertOnSubmit(j);
                    context.SubmitChanges();
                }
            }

                //XmlDocument xml = new XmlDocument();
                //xml.Load(attach.ContentStream);
        }
    }
}
