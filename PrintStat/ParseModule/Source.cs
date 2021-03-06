﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


using System.Xml;
using System.Xml.Linq;
using System.IO.Compression;
using System.Globalization;

using S22.Imap;
using System.Net.Mail;

namespace PrintStat.ParseModule
{
    public class Source
    {
        public Source(KeyValuePair<uint, MailMessage>  mes)
        {
            context = new PrintStatDataDataContext();
            message = mes;        
            
        }
        public KeyValuePair<uint, MailMessage>  message;
        //public TempJob job;
        public PrintStatDataDataContext context;
        public int UPID;
        public int mID;
        public string GetAtribute(XmlElement x, string tag, string attName)
        {
            string result = "";
            foreach (XmlAttribute att in x.GetElementsByTagName(tag)[0].Attributes)
            {
                if (att.Name == attName)
                {
                    result = att.Value;
                    break;
                }
            }
            return result;
        }
        public string GetInnerText(XmlElement x, string tag)
        {
            return x.GetElementsByTagName(tag)[0].InnerText;
        }
        public virtual void GetValueTag(Tag t, Job j, XmlElement x)
        {

        }
        public virtual void parce()
        {
       
        }
    }
}