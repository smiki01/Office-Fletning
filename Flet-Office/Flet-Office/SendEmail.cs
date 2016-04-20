using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace Flet_Office
{
    public class SendEmail
    {
        public SendEmail(string MedlemsNr, string IntNavn, string AfdMail, string FilePDF, string emailModt, string MailLog, string Kategori)
        {
            Chilkat.MailMan mailman = new Chilkat.MailMan();

            bool success = mailman.UnlockComponent("PROBITMAILQ_ckTyt6oroW8J");
            if (success != true)
            {
                MyVars.EmailSendtOK = false;
                MyVars.MailFejltekst = "Licens problemer: " + mailman.LastErrorText;
                MyVars.emailBody = "";
                MyVars.emailSubj = "";
                return;
            }
            mailman.SmtpHost = "172.16.2.150";
            //  Set the SMTP login/password (if required)
            //mailman.SmtpUsername = "Plaza.local\\sa-batch";
            //mailman.SmtpPassword = "8RUG2plaza";
            Chilkat.Email email = new Chilkat.Email();

            if(MedlemsNr != "")
            {
                if(Kategori != "")
                {
                    EGBoligData EGData = new EGBoligData();
                    Mail oMailData = EGData.GetMailData(Kategori);       // her hentes mail-tekster fra EGBolig tabel
                    if (MyVars.emailSubj == "")
                    {
                        MyVars.emailSubj = oMailData.Emne;
                    }
                    MyVars.emailSubj = MyVars.emailSubj.Replace("[Selskab_Nr]", "099");
                    MyVars.emailSubj = MyVars.emailSubj.Replace("[Medlem_Nr]", MedlemsNr);
                    string emailBodyx = oMailData.Tekst;
                    
                    emailBodyx = emailBodyx.Replace("[Medlem_Navn]", IntNavn);
                    emailBodyx = emailBodyx.Replace("[Bruger_Navn]", "");
                    if (MyVars.emailBody == "")
                    {
                        MyVars.emailBody += emailBodyx;
                    }
                }
                else
                {
                    MyVars.emailSubj = "Medlemsbrev fra Boligkontoret Danmark";
                    MyVars.emailBody = "Medlemsbrev fra Boligkontoret Danmark";
                }
            }

            if (Properties.Settings.Default.Common_PROD == false)
            {
                emailModt = Properties.Settings.Default.Common_Email_Testmodt;    //her sikres at alle emails sendes til kism (anvendes til test)
            }

            email.Subject = MyVars.emailSubj;
            email.Body = MyVars.emailBody;
            email.From = AfdMail;
      //emailModt = "kism@bdk.dk";
            email.AddBcc("", emailModt);

            if (FilePDF != "")
            {
                string contentType = email.AddFileAttachment(FilePDF);
                if (string.IsNullOrEmpty(contentType))
                {
                    MyVars.EmailSendtOK = false;
                    MyVars.MailFejltekst = "Dokument kunne ikke vedhæftes e-mail pga. fejl: " + mailman.LastErrorText;
                    MyVars.emailBody = "";
                    MyVars.emailSubj = "";
                    return;
                }
            }

            success = mailman.SendEmail(email);
            if (success != true)
            {
                MyVars.EmailSendtOK = false;
                MyVars.emailBody = "";
                MyVars.emailSubj = "";
                return;
            }
            else
            {
                MyVars.EmailSendtOK = true;
                MyVars.MailFejltekst = "";
            }

            success = mailman.CloseSmtpConnection();
            if (success != true)
            {
                MyVars.EmailSendtOK = false;
                MyVars.MailFejltekst = "Connection fejl opstået ved afsendelse af e-mail: Connection to SMTP server not closed cleanly: " + mailman.LastErrorText;
            }
            MyVars.emailBody = "";
            MyVars.emailSubj = "";
        }
    }
}
