using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using S22.Imap;
using System.Net.Mail;

namespace PrintStat.ParseModule
{
    public class GetData
    {

        private Email em;

        private Dictionary<uint, MailMessage> mes;
        private Source source;

        public GetData(DateTime fromDate, DateTime ToDate)
        {

            var context = new PrintStatDataDataContext();
            string server = context.SettingValue.First(p => p.Profile.Name == "Email" && p.Settings.Name == "Сервер").Value,//"tpnic.ru",
            port = context.SettingValue.First(p => p.Profile.Name == "Email" && p.Settings.Name == "Порт").Value,//"143",
            user = context.SettingValue.First(p => p.Profile.Name == "Email" && p.Settings.Name == "E-mail адрес").Value,//"printer.stats@tpnic.ru",
            pass = context.SettingValue.First(p => p.Profile.Name == "Email" && p.Settings.Name == "Пароль").Value;//"printer";

           
            em = new Email(server, int.Parse(port), user, pass,fromDate,ToDate, false, "INBOX");
            mes = em.GetMessages();
            ////
        }

       public void GetSource()
        {
            foreach (var k in mes)
            {
                //listBox1.Items.Add(k.Value.Subject);
                if (k.Value.Subject == "Accounting report. Printer Serial Number: CN314DH01B")
                {
                    source = new SourceHPXml();
                    source.parce(k);
                    //ParseXmlHP(k);
                }
                if (k.Value.Subject == "FS-C8600DN log")
                {

                    source = new SourceFSXml();
                    source.parce(k);
                    //ParseXmlFS(k);
                }

            }
        }
    }
}