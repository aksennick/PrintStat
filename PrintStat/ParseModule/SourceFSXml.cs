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

        public SourceFSXml(KeyValuePair<uint, MailMessage> mes)
            : base(mes)
        {
            //UPID = context.Device.FirstOrDefault(p => p.SearchString == "hp").ID;
            mID = context.Device.FirstOrDefault(p => p.SearchString == "FS").ModelID;
        }
        public override void GetValueTag(Tag t, Job j, XmlElement x)
        {
            j.DeviceID = context.Device.FirstOrDefault(p => p.SearchString == "fs").ID;
            switch (t.Name)
            {
                case "Задание":
                    {
                        j.Name = GetInnerText(x, t.Tag1);
                        break;
                    }
                case "Страницы":
                    {
                        j.Pages = Convert.ToInt32(GetInnerText(x, t.Tag1));
                        break;
                    }
                case "Копии":
                    {
                        j.Copies = Convert.ToInt32(GetInnerText(x, t.Tag1));
                        break;
                    }
                case "Время начала":
                    {
                        j.StartTime = new DateTime(Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:year")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:month")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:day")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:hour")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:minute")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:second")));
                       break;
                    }
                case "Время окончания":
                    {
                        j.EndTime = new DateTime(Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:year")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:month")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:day")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:hour")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:minute")),
                                                    Convert.ToInt32(GetInnerText((XmlElement)x.GetElementsByTagName(t.Tag1)[0], "kmloginfo:second")));
                        break;
                    }                      
                
                case "Выполнил":
                    {
                        try
                        {
                            j.UserTabNumber = context.Employee.First(p => p.TabNumber == GetInnerText(x, t.Tag1)).TabNumber;
                        }
                        catch
                        {
                            j.UserTabNumber = "1369";
                        }
                        break;
                    }
            }

        }

        public override void parce()
        {
            var mt = context.ModelTag.Where(t => t.ModelID == mID);
            foreach (var attach in message.Value.Attachments)
            {XmlDocument xml = new XmlDocument();

                xml.Load(attach.ContentStream);

                foreach (XmlElement x in xml.GetElementsByTagName("kmloginfo:print_job_log"))
                {
                        var j = new Job();

                        foreach (ModelTag tag in mt)
                        {
                            GetValueTag(tag.Tag, j, x);
                        }
                        j.Duration = Convert.ToInt32((j.StartTime.Value - j.StartTime.Value).TotalMinutes);
                        j.ApplicationID = context.Application.FirstOrDefault(p => p.Name == "Default").ID;
                        context.Job.InsertOnSubmit(j);
                        context.SubmitChanges();
                    
                }
            }
        }
    }
}
