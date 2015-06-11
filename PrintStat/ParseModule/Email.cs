using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using S22.Imap;
using System.Net.Mail;

namespace PrintStat.ParseModule
{

    public class Email
    {

        private static ImapClient imapClient = null;
        private static IEnumerable<string> mailboxes;
        private static IEnumerable<uint> mailIDs;
        private static Dictionary<uint, MailMessage> mailMessages;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="host">Host Address</param>
        /// <param name="port">Port SSL or Not</param>
        /// <param name="user">Username i.e. email</param>
        /// <param name="pass">Password</param>
        /// <param name="SSL">true/false to SSL connection</param>
        /// <param name="mailbox">Default mailbox</param>
        public Email(string host, int port, string user, string pass, DateTime fromDate, DateTime ToDate, bool SSL = false, string mailbox = "INBOX")
        {
            //imapClient = new ImapClient(host, port, SSL);
            imapClient = new ImapClient(host, port, SSL);
            imapClient.Login(user, pass, AuthMethod.Login);// AuthMethod.Auto
            mailboxes = imapClient.ListMailboxes();
            GetMail(fromDate, ToDate, mailbox);
        }

        /// <summary>
        /// Private get mail. Populates dictionary
        /// </summary>
        /// <param name="mailbox"></param>
        private void GetMail(DateTime fromDate, DateTime ToDate, string mailbox = "INBOX")
        {
            //var t = new DateTime(2015, 03, 25);
            //var sc = SearchCondition.Before(t);

            // var sc = new SearchCondition().Not(SearchCondition.Before(t));
            var SC = SearchCondition.Before(fromDate);
            var NowSC = SearchCondition.Before(ToDate);
            var search = SC.Not(NowSC);
            mailMessages = new Dictionary<uint, MailMessage>();
            if (isAuthed())
                mailIDs = imapClient.Search(search, mailbox);//Unseen()(new SearchCondition()).Not(
            foreach (var id in mailIDs)
            {
                MailMessage mail = imapClient.GetMessage(id, FetchOptions.Normal, false);

                mail.BodyEncoding = Encoding.Unicode;
                mailMessages.Add(id, mail);
            }
        }

        /// <summary>
        /// Logout of account
        /// </summary>
        public void LogOut()
        {
            if (imapClient.Authed && imapClient != null)
                imapClient.Logout();
            imapClient.Dispose();
        }

        /// <summary>
        /// Is authorised to use account
        /// </summary>
        /// <returns>boolean</returns>
        public bool isAuthed()
        {
            return imapClient.Authed;
        }

        /// <summary>
        /// Total mail count
        /// </summary>
        /// <returns>integer representation</returns>
        public int MailCount()
        {
            int i = 0;
            if (isAuthed())
                foreach (var item in mailIDs)
                {
                    i++;
                }
            return i;
        }

        /// <summary>
        /// Returns mail messages - key/value dictionary. UID of email is key. .NET MailMessage is value
        /// </summary>
        /// <returns>Dictionary<uid,MailMessage></returns>
        public Dictionary<uint, MailMessage> GetMessages()
        {
            if (isAuthed())
                return mailMessages;
            return null;
        }

        /// <summary>
        /// Move email from one location to another
        /// </summary>
        /// <param name="uid">Email ID - found in GetMessages class. Key in dicionary</param>
        /// <param name="mailbox">Mailbox to move email to</param>
        /// <returns>Boolean if successful</returns>
        public bool MailMove(uint uid, string mailbox = "Trash")
        {
            if (isAuthed())
            {
                imapClient.MoveMessage(uid, mailbox);
                //Expunge finalises delete.
                imapClient.Expunge();
                return true;
            }
            return false;
        }


    }
}
