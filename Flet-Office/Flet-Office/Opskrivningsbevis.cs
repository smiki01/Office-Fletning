using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.IO;
using System.Threading.Tasks;
using System.EnterpriseServices;
using Microsoft.Office.Interop.Word;

namespace Flet_Office
{
    public partial class Opskrivningsbevis : Form
    {
        Medlemskab Medlem;
        SelskabSet Selskab;
        InteressentSet Interessent;
        string AfdKontorNavn = "";
        string AfdKontorAdresse = "";
        string AfdKontorPostby = "";
        string AfdKontorTlf = "";
        string AfdKontorFax = "";
        string AfdKontorEmail = "";
        string AfdKtTrafTid1 = "";
        string AfdKtTrafTid2 = "";
        string AfdKtTrafTid3 = "";
        string AfdKtTrafTid4 = "";
        string AfdKtTrafTid5 = "";
        string FirmaNavn = "";
        string IntNavn = "";
        string IntNr = "";
        string IntCo = "";
        string IntAdresse = "";
        string IntLokalBy = "";
        string IntPostBy = "";
        string IntInternetNr = "";
        string IntPassword = "";
        string IntGobUrl = "";
        string PostStregkode = "";
        string SelskabsNr = "";
        string gemSelskab = "";
        string bookSave = "";
        string gemMedl = "";
        string SelNavn = "";
        string AfdMail = "";
        string LokalBy = "";
        string KontorNavn = "";
        string emailModt = "";
        bool bonus_Sel = false;
        string fejlTekst = "";
        string GoBib = Properties.Settings.Default.Common_GoDokBib;
        string PgmLog = "";
        string PgmLogTest = "";
        string PgmLogSys = "";
        string LogText = "";
        int FejlCnt = 0;
        int MedlCnt = 0;
        string outputFolderPath = Properties.Settings.Default.Opskrivningsbevis_Folder;
        string FilePDF = ""; 

        public Opskrivningsbevis()
        {
            InitializeComponent();
            if (MyVars.ParmYes == "AUTO")
            {
                this.timer1.Enabled = true;
                this.timer1.Interval = 5000;
                this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            }
            Medlem = new Medlemskab("");                              
            string skabelonPath = Properties.Settings.Default.Common_SkabelonSti;
            string skabelonOriginal = Properties.Settings.Default.Opskrivningsbevis_SkabelonNavn;
            string outputFolderOld = Properties.Settings.Default.Opskrivningsbevis_FolderOld;
            string dt = DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss");
            PgmLog = Properties.Settings.Default.Opskrivningsbevis_ErrorLog + dt + ".Log";
            PgmLogSys = "C:\\FletOffice\\ProgramError_Opskriv_SysLog_" + dt + ".Log";
            PgmLogTest = "C:\\FletOffice\\ProgramError_Opskriv_TestLog_" + dt + ".Log";

            if (Directory.Exists(outputFolderOld))
            {
                System.IO.Directory.Delete(outputFolderOld, true);
            }       
            System.IO.Directory.Move(outputFolderPath, outputFolderOld);
            System.IO.Directory.CreateDirectory(outputFolderPath);

            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();

            int cnt = 0;

            foreach (Medlem oMedlemsData in Medlem.AlleMedl)      //her løbes alle dagens nye betalende medlemmer
            {
                //her behandles medlemsdata
                Document skabelon = application.Documents.Open(skabelonPath + skabelonOriginal);
                application.Visible = false;
                if (gemMedl != oMedlemsData.MedlemsNr && bookSave != "" && gemMedl != "")
                {
                    //Her behandles multiline bogmærker, PDF generering, Email afsendelse og journalisering til Gobolig
                    BehandlMedlem(skabelon, gemMedl);

                    Document document = application.Documents.Open(skabelonPath + skabelonOriginal);
                    application.Visible = false;

                    // clone the doc object.
                    skabelon = document;
                }
                else
                {
                    if (gemMedl != oMedlemsData.MedlemsNr && bookSave == "" && gemMedl != "")
                    {
                        if (bonus_Sel == true)
                        {
                            fejlTekst = ", som er en tilføjelse af bonusselskab og derfor ikke skal have tilsendt et opskrivningsbevis.";
                        }
                        else
                        {
                            fejlTekst = ", som kun er medlem af selskab 99 og derfor ikke får et opskrivningsbevis tilsendt.";
                        }
                        textBox1.Text += "Nu har vi behandlet medlem nr.: " + gemMedl + fejlTekst + Environment.NewLine;
                        LogText = "Nu har vi behandlet medlem nr.: " + gemMedl + fejlTekst;
                        File.AppendAllText(PgmLog, LogText + Environment.NewLine);
                        MyVars.Selopdat = "";
                        // her skal Word dokument slettes og hoppes videre til næste medlem.
                        ((_Document)skabelon).Close(WdSaveOptions.wdDoNotSaveChanges);   //her lukkes skabelonen
                        //File.AppendAllText(PgmLogTest, "Skabelon lukkes og ny åbnes..." + Environment.NewLine);
                        Document document = application.Documents.Open(skabelonPath + skabelonOriginal);
                        application.Visible = false;

                        // clone the doc object.
                        skabelon = document;
                    }
                }
                //her hentes selskabsdata
                SelNavn = "";
                UdtrSelOpl(oMedlemsData.ListeMedlemskab);

                //her hentes interessentdata
                UdtrIntOpl(oMedlemsData.MedlemsNr, KontorNavn);

                //her hentes interessentdata for kolofon
                UdtrKolOpl(oMedlemsData.MedlemsNr, "Hovedkontoret");

                //her loopes igennem alle bookmarks i dokumentet.
                OpdaterBookmark(skabelon, oMedlemsData.MedlemsNr, oMedlemsData.ListeMedlemskab, oMedlemsData.DatoOpnot, oMedlemsData.Status);
                cnt++;

                gemMedl = oMedlemsData.MedlemsNr;
                if (cnt == Medlem.AntalMedl)                // her behandles sidste medlem i SQL udtrækket færdigt
                {
                    //Her behandles multiline bogmærker, PDF generering, Email afsendelse, journalisering til Gobolig og lukning af skabelon.
                    BehandlMedlem(skabelon, gemMedl);
                }
                MyVars.InsertOk = false;

            }
            //Her afsluttes programmet.....
            string Server = Properties.Settings.Default.Common_MonitorServerNavn;
            string Job = Properties.Settings.Default.Opskrivningsbevis_MonitorJobNavn;
            string ServerGruppe = Properties.Settings.Default.Common_MonitorGruppeNavnGO;
            string Tekst = "";
            int Alarm = Properties.Settings.Default.Common_MonitorAlarmVærdiOK;
            int antMailSendt = MedlCnt - FejlCnt;

            if (File.Exists(PgmLogSys))
            {
                MyVars.emailSubj = "Systemfejllog for Opskrivningsbreve";
                MyVars.emailBody = "Dette er en fejllog som viser system relaterede fejl ved kørsel af opskrivningsbreve. " + Environment.NewLine + "Se vedhæftede log for detaljer";
                SendEmail EmailObj = new SendEmail("", "", "admin@bdk.dk", PgmLogSys, "sa-gobolig@bdk.dk", "");
            }

            if (File.Exists(PgmLogTest))
            {
                MyVars.emailSubj = "Testlog for Opskrivningsbevis";
                MyVars.emailBody = "Dette er en testlog som anvendes ved test af Opskrivningsbevis. " + Environment.NewLine + "Se vedhæftede log for detaljer";
                SendEmail EmailObj = new SendEmail("", "", "admin@bdk.dk", PgmLogTest, "kism@bdk.dk", "");
            }

            if (File.Exists(PgmLog))
            {
                string afs = Properties.Settings.Default.Common_EmailAfsFejl;
                string modt = Properties.Settings.Default.Common_EmailEGModtLog;

                if (FejlCnt > 0)
                {
                    MyVars.emailSubj = Properties.Settings.Default.Opskrivningsbevis_EmailSubjFejl;
                    MyVars.emailBody = Properties.Settings.Default.Common_EmailBodyFejl +" (Opskrivningsbevis) " + Environment.NewLine + "Se vedhæftede log for detaljer";
                    Tekst = Properties.Settings.Default.Common_MonitorTekstFejl + FejlCnt + ". " + Properties.Settings.Default.Common_MonitorTekst + antMailSendt;
                    Alarm = Properties.Settings.Default.Common_MonitorAlarmVærdiFejl;
                }
                else
                {
                    MyVars.emailBody = Properties.Settings.Default.Common_EmailBodyOk + " (Opskrivningsbevis) " + Environment.NewLine + "Se vedhæftede log for detaljer";
                    MyVars.emailSubj = "Opskrivningsbevis: " +Properties.Settings.Default.Common_EmailSubjOk;
                    Tekst = Properties.Settings.Default.Common_MonitorTekstOK + ". " + Properties.Settings.Default.Common_MonitorTekst + antMailSendt;
                }
                LogText = Tekst;
                File.AppendAllText(PgmLog, LogText + Environment.NewLine);

                SendEmail EmailObj = new SendEmail("", "", afs, PgmLog, modt, "");
            }
            else
            {
                Tekst = Properties.Settings.Default.Common_MonitorTekstOK + ". " + Properties.Settings.Default.Common_MonitorTekst + antMailSendt;
            }

            if (MedlCnt == 0)
            {
                textBox1.Text += Properties.Settings.Default.Opskrivningsbeviser_SlutTekstIngen + Environment.NewLine;
                Tekst = Properties.Settings.Default.Opskrivningsbeviser_SlutTekstIngen + Environment.NewLine;
            }

            MonitorAdd MonitorAddGO = new MonitorAdd(Server, Job, Tekst, Alarm, ServerGruppe);
            ServerGruppe = Properties.Settings.Default.Common_MonitorGruppeNavnEG;
            MonitorAdd MonitorAddEG = new MonitorAdd(Server, Job, Tekst, Alarm, ServerGruppe);

            // Close word.
            ((_Application)application).Quit();                                  //her lukkes Word applicationen
        }
   


        //Herunder følger alle funktionerne som kaldes fra hovedprogrammet
        public void BehandlMedlem(Document skabelon, string MedlemsNr)
        {
            //her opdateres multiline bogmærkerne 
            SkrivBookmark(skabelon, bookSave);

            textBox1.Text += Properties.Settings.Default.Common_SlutTekst + gemMedl + ". (INT-" + IntNr.PadLeft(8, '0') + ")" + Environment.NewLine;

            //her konverteres dokumentet til PDF
            FilePDF = @outputFolderPath + "\\OpskrivningsBevis_MedlemNr_" + gemMedl + ".pdf";
            PDF PDFObj = new PDF(skabelon, FilePDF);

            //her vedhæftes det udfyldte PDF dokument til en email og derefter slettes det og mail sendes, herefter genoptages loop (break)
            UdtrKolOpl(gemMedl, LokalBy);          //her hentes afdelingskontors email adresse (til afsender)
            if (AfdKontorEmail == "" || AfdKontorEmail.IndexOf("Afdelingskontoret") == -1)
            {
                AfdKontorEmail = AfdMail;
            }
            string Kategori = "0002009999";
            SendEmail EmailObj = new SendEmail(gemMedl, IntNavn, AfdKontorEmail, FilePDF, emailModt, Kategori);

            if (MyVars.EmailSendtOK == false)
            {
                string gemEmailBody = MyVars.emailBody;
                if (MyVars.MailFejltekst != "")
                {
                    MyVars.emailSubj = "System fejl opstået ved send af e-mail!";
                    MyVars.emailBody = "Herunder kan læses fejlmeddelelse fra send e-mail program: " + Environment.NewLine + MyVars.MailFejltekst;
                    LogText = "System fejl opstået ved send af e-mail til medlem " + gemMedl + " på følgende Email adresse (" + emailModt + ")." + Environment.NewLine + "Fejltekst: " + MyVars.MailFejltekst;
                    File.AppendAllText(PgmLogSys, LogText);
                    FejlCnt++;
                }
                else
                {
                    MyVars.emailSubj = Properties.Settings.Default.Common_EmailEjLev + gemMedl + ". Handling påkrævet!";
                    MyVars.emailBody = "Nedenstående email kunne ikke leveres til medlemsnr. 099-" + gemMedl + "." + Environment.NewLine;
                    if (emailModt != "")
                    {
                        MyVars.emailBody += "Årsag: Email adresse (" + emailModt + ") er ikke gyldig.";
                    }
                    else
                    {
                        MyVars.emailBody += "Årsag: " + Properties.Settings.Default.Common_EmailIngen + " " + IntNavn + Environment.NewLine;
                    }
                    MyVars.emailBody += Environment.NewLine + "Forsøg evt. at sende email igen fra interessenten i Gobolig (INT-" +
                        IntNr.PadLeft(8, '0') + ")." + Environment.NewLine + "Alternativt, send opskrivningsbevis pr. brev i stedet for til:" + Environment.NewLine +
                        IntNavn + Environment.NewLine + IntAdresse + Environment.NewLine + IntPostBy +
                        Environment.NewLine + Environment.NewLine + Environment.NewLine + gemEmailBody;
                    textBox1.Text += Properties.Settings.Default.Common_EmailEjLev + " 099-" + gemMedl + ". Email adresse (" + emailModt + ") er ikke gyldig. " + Environment.NewLine;
                    if (emailModt != "")
                    {
                        LogText = Properties.Settings.Default.Common_EmailEjLev + gemMedl + ". Email adresse (" + emailModt + ") er ikke gyldig." + Environment.NewLine;
                    }
                    else
                    {
                        LogText = Properties.Settings.Default.Common_EmailIngen + " " + IntNavn + " (99-" + gemMedl + "). Opskrivningsbevis sendt til printer.";
                    }

                    File.AppendAllText(PgmLog, LogText + Environment.NewLine);

                    //Hvis e-mail ikke kunne sendes til medlem, sendes vedhæftede PDF dokument til default printer 
                    PDFPrint.PrintPDFs(FilePDF);

                    FejlCnt++;
                }
            }
            else
            {
                int count = MedlCnt + 1;
                LogText = count + ". Opskrivningsbevis sendt ok på e-mail til medlem: " + IntNavn + " (99-" + gemMedl + ") på denne e-mail adresse: " + emailModt + Environment.NewLine;
                File.AppendAllText(PgmLog, LogText);
            }

            //her tjekkes om interessent allerede findes i Gobolig
            int Iintval = 30;
            int Iantloop = 6 * Iintval;
            int Isleep = Iintval * 1000;
            int ICnt = 0;
            while (IntGobUrl == "" && ICnt < Iantloop)
            {
                //her indsættes interessent oprettelses entry ind i ESDHMJFejl tabellen
                if (MyVars.InsertOk == false)
                {
                    int IntNummer = Convert.ToInt32(IntNr);
                    string PersId = "INT|||kism¤INSERT¤" + IntNummer;
                    ESDHMJFejlAdd oESDHMJFejlAdd = new ESDHMJFejlAdd(IntNummer, PersId);
                }

                System.Threading.Thread.Sleep(Isleep);       //vent 30 sekunder og tjek herefter om interessent er oprettet (max 6 gange=2 min.)
                UdtrIntOpl(MedlemsNr, KontorNavn);

                ICnt = ICnt + Iintval;
            }

            //her journaliseres dokument til Gobolig
            if (IntGobUrl != "")                            //journalisering sker hvis der findes en gobUrl for interessenten
            {
                Gobolig GoboligObj = new Gobolig(IntGobUrl, FilePDF, GoBib);
                if (MyVars.GoJournalOk == false)
                {
                    textBox1.Text += Properties.Settings.Default.Common_GoboligJourFejl + gemMedl + Environment.NewLine;
                    LogText = Properties.Settings.Default.Common_GoboligJourFejl + gemMedl;
                    File.AppendAllText(PgmLog, LogText + Environment.NewLine);
                    FejlCnt++;
                }
            }
            else
            {
                textBox1.Text += Properties.Settings.Default.Common_GoboligJourFejl + gemMedl + " fordi interessentnr = INT-" 
                    + IntNr.PadLeft(8, '0') + " ikke findes i Gobolig på journaliseringstidspunktet." + Environment.NewLine;
                LogText = Properties.Settings.Default.Common_GoboligJourFejl + gemMedl + " fordi interessentnr = INT-"
                    + IntNr.PadLeft(8, '0') + " ikke findes i Gobolig på journaliseringstidspunktet.";
                File.AppendAllText(PgmLog, LogText + Environment.NewLine);
                FejlCnt++;
            }

            //her lukkes skabelonen
            //File.AppendAllText(PgmLogTest, "Skabelon lukkes herefter.." + Environment.NewLine);
            ((_Document)skabelon).Close(WdSaveOptions.wdDoNotSaveChanges);   //her lukkes skabelonen
            MedlCnt++;
            gemMedl = "";
            gemSelskab = "";
            bookSave = "";
        }


        public void SkrivBookmark(Document skabelon, string bookSave)
        {
            foreach (Bookmark bookmark in skabelon.Bookmarks)
            {
                //her isættes funden tekst for flere liniers bogmærker
                string bMark = bookmark.Name;
                switch (bMark)
                {
                    case "Medlem_Liste_MedlemSelskab":
                        bookmark.Range.Font.Size = 12.0F;
                        bookmark.Range.Bold = 2;
                        bookmark.Range.Text = bookSave;
                        break;
                    case "Medlem_Liste_NSelskab":
                        bookmark.Range.Text = "";
                        break;
                    case "Medlem_Liste_Opnoteringsdato":
                        bookmark.Range.Text = "";
                        break;
                    case "Medlem_EmailAdr":
                        if (emailModt != "")
                        {
                            bookmark.Range.Text = emailModt;
                        }
                        else
                        {
                            bookmark.Range.Text = "Ingen e-mail adresse oplyst.";
                        }
                        break;
                    default:
                        break;
                }
            }
        }
        public void UdtrIntOpl(string MedlemNr, string KontorNavn)
        {
            Interessent = new InteressentSet(Convert.ToInt32(MedlemNr), KontorNavn);
            foreach (Interessent oInteressentData in Interessent.AlleInt)
            {
                IntNavn = oInteressentData.IntNavn;
                IntNr = oInteressentData.InteressentNr;
                IntCo = oInteressentData.IntCO;
                IntAdresse = oInteressentData.IntAdresse;
                IntLokalBy = oInteressentData.IntLokalBy;
                IntPostBy = oInteressentData.IntPostBy;
                IntInternetNr = oInteressentData.IntInternetNr;
                IntPassword = oInteressentData.IntPassword;
                IntGobUrl = oInteressentData.IntGobUrl;
                emailModt = oInteressentData.IntEmail;
                PostStregkode = oInteressentData.PostStregkode;
            }
  //emailModt = "smiki01@hotmail.dk";          //kun til test brug - invalid e-mail adresse
  //emailModt = "";                            //kun til test brug - manglende e-mail adresse
        }
        public void UdtrKolOpl(string MedlemNr, string KontorNavn)
        {
            Interessent = new InteressentSet(Convert.ToInt32(MedlemNr), KontorNavn);
            foreach (Interessent oInteressentData in Interessent.AlleInt)
            {
                AfdKontorNavn = oInteressentData.IntNavn;
                AfdKontorAdresse = oInteressentData.IntAdresse;
                AfdKontorPostby = oInteressentData.IntPostBy;
                AfdKontorTlf = oInteressentData.IntTlf;
                AfdKontorFax = oInteressentData.IntFax;
                AfdKontorEmail = oInteressentData.IntEmail;
                AfdKtTrafTid1 = oInteressentData.IntTraftid1;
                AfdKtTrafTid2 = oInteressentData.IntTraftid2;
                AfdKtTrafTid3 = oInteressentData.IntTraftid3;
                AfdKtTrafTid4 = oInteressentData.IntTraftid4;
                AfdKtTrafTid5 = oInteressentData.IntTraftid5;
            }
        }
        public void UdtrSelOpl(string ListeMedlemskab)
        {
            Selskab = new SelskabSet(Convert.ToInt32(ListeMedlemskab));

            foreach (Selskab oSelskabsData in Selskab.AlleSel)
            {
                SelNavn = oSelskabsData.SelNavn;
                AfdMail = oSelskabsData.AfdMail;
                LokalBy = oSelskabsData.LokalBy;
            }
            SelskabsNr = ListeMedlemskab;
            Selskab = new SelskabSet(99);
            foreach (Selskab oSelskabsData in Selskab.AlleSel)
            {
                FirmaNavn = oSelskabsData.SelNavn;
            }
        }

        public void OpdaterBookmark(Document skabelon, string MedlemsNr, string ListeMedlemskab, string DatoOpnot, string Status)
        {
            foreach (Bookmark bookmark in skabelon.Bookmarks)
            {
                //her isættes funden tekst for bogmærker pr. medlem
                string bMark = bookmark.Name;
                switch (bMark)
                {
                    case "Adresse_Navn_AfdKontor":
                        bookmark.Range.Font.Size = 7.50F;
                        bookmark.Range.Bold = 2;
                        bookmark.Range.Text = AfdKontorNavn;
                        break;
                    case "Adresse_Adresse_AfdKontor":
                        bookmark.Range.Font.Size = 7.50F;
                        bookmark.Range.Text = AfdKontorAdresse;
                        break;
                    case "Adresse_Postby_AfdKontor":
                        bookmark.Range.Font.Size = 7.50F;
                        bookmark.Range.Text = AfdKontorPostby;
                        break;
                    case "Adresse_Tlf1_AfdKontor":
                        bookmark.Range.Font.Size = 7.50F;
                        bookmark.Range.Text = AfdKontorTlf;
                        break;
                    case "Adresse_Email1_AfdKontor":
                        bookmark.Range.Font.Size = 7.50F;
                        bookmark.Range.Text = AfdKontorEmail;
                        break;
                    case "Adresse_Telefon_Fax_AfdKontor":
                        bookmark.Range.Font.Size = 7.50F;
                        bookmark.Range.Text = AfdKontorFax;
                        break;
                    case "Adresse_Træffetid1_AfdKontor":
                        bookmark.Range.Font.Size = 7.50F;
                        bookmark.Range.Text = AfdKtTrafTid1;
                        break;
                    case "Adresse_Træffetid2_AfdKontor":
                        bookmark.Range.Font.Size = 7.50F;
                        bookmark.Range.Text = AfdKtTrafTid2;
                        break;
                    case "Adresse_Træffetid3_AfdKontor":
                        bookmark.Range.Font.Size = 7.50F;
                        bookmark.Range.Text = AfdKtTrafTid3;
                        break;
                    case "Adresse_Træffetid4_AfdKontor":
                        bookmark.Range.Font.Size = 7.50F;
                        bookmark.Range.Text = AfdKtTrafTid4;
                        break;
                    case "Adresse_Træffetid5_AfdKontor":
                        bookmark.Range.Font.Size = 7.50F;
                        bookmark.Range.Text = AfdKtTrafTid5;
                        break;
                    case "Medlem_Nr":
                        bookmark.Range.Font.Size = 12.0F;
                        bookmark.Range.Bold = 2;
                        bookmark.Range.Text = MedlemsNr;
                        break;
                    case "Medlem_Selskab":
                        bookmark.Range.Font.Size = 12.0F;
                        bookmark.Range.Bold = 2;
                        bookmark.Range.Text = "099";
                        break;
                    case "Selskab_Navn2":
                        bookmark.Range.Text = "";
                        break;
                    case "Firma_Navn":
                        bookmark.Range.Text = FirmaNavn;
                        break;
                    case "Firma_Navn2":
                        bookmark.Range.Text = FirmaNavn;
                        break;
                    case "Firma_Navn3":
                        bookmark.Range.Text = FirmaNavn;
                        break;
                    case "Medlem_Navn":
                        bookmark.Range.Text = IntNavn;
                        break;
                    case "Medlem_Navn2":
                        bookmark.Range.Text = IntNavn;
                        break;
                    case "Medlem_CO":
                        bookmark.Range.Text = IntCo;
                        break;
                    case "Medlem_CO2":
                        bookmark.Range.Text = IntCo;
                        break;
                    case "Medlem_Adresse":
                        bookmark.Range.Text = IntAdresse;
                        break;
                    case "Medlem_Adresse2":
                        bookmark.Range.Text = IntAdresse;
                        break;
                    case "Medlem_Lokalby":
                        bookmark.Range.Text = IntLokalBy;
                        break;
                    case "Medlem_Lokalby2":
                        bookmark.Range.Text = IntLokalBy;
                        break;
                    case "Medlem_Postby":
                        bookmark.Range.Text = IntPostBy;
                        break;
                    case "Medlem_Postby2":
                        bookmark.Range.Text = IntPostBy;
                        break;
                    case "Medlem_Internetnr":
                        bookmark.Range.Font.Size = 12.0F;
                        bookmark.Range.Bold = 2;
                        if (IntInternetNr == "")
                        {
                            IntInternetNr = IntNr;
                        }
                        bookmark.Range.Text = IntInternetNr;
                        break;
                    case "Dato_side1":
                        bookmark.Range.Font.Size = 9.50F;
                        bookmark.Range.Text = DateTime.Today.ToString("dd. MMMM yyyy");
                        break;
                    case "Dato_side2":
                        bookmark.Range.Font.Size = 9.50F;
                        bookmark.Range.Text = DateTime.Today.ToString("dd. MMMM yyyy");
                        break;
                    case "Post_Stregkode":
                        bookmark.Range.Font.Size = 9.50F;
                        bookmark.Range.Bold = 2;
                        bookmark.Range.Text = PostStregkode;
                        break;
                    case "Medlem_Liste_MedlemSelskab":
                        if(ListeMedlemskab != "99")
                        {
                            bonus_Sel = true;
                        }
                        else
                        {
                            bonus_Sel = false;
                        }
                        if (ListeMedlemskab != "99" && gemSelskab != SelskabsNr && SelNavn != "")
                        {
                            bookmark.Range.Font.Size = 12.0F;
                            bookmark.Range.Bold = 2;
                            if ((gemMedl == MedlemsNr && (Status == "0" || Status == "1" || Status == "8" || Status == "9")) || gemMedl == "" ) 
                            {
                                if (bookmark.Name == "Medlem_Liste_MedlemSelskab")
                                {
                                    bookSave += ListeMedlemskab;
                                }
                            }
                        }
                        break;
                    case "Medlem_Liste_NSelskab":
                        if (ListeMedlemskab != "99" && gemSelskab != SelskabsNr && SelNavn != "")
                        {
                            bookmark.Range.Font.Size = 12.0F;
                            bookmark.Range.Bold = 2;
                            if ((gemMedl == MedlemsNr && (Status == "0" || Status == "1" || Status == "8" || Status == "9")) || gemMedl == "")
                            {
                                if (bookmark.Name == "Medlem_Liste_NSelskab")
                                {
                                    bookSave += " " + SelNavn;
                                }
                            }
                        }
                        break;
                    case "Medlem_Liste_Opnoteringsdato":
                        if (ListeMedlemskab != "99" && gemSelskab != SelskabsNr && SelNavn != "")
                        {
                            bookmark.Range.Font.Size = 12.0F;
                            bookmark.Range.Bold = 2;
                            if ((gemMedl == MedlemsNr && (Status == "0" || Status == "1" || Status == "8" || Status == "9")) || gemMedl == "")
                            {
                                if (bookmark.Name == "Medlem_Liste_Opnoteringsdato")
                                {
                                    if (DatoOpnot != "")
                                    {
                                        bookSave += " \t" + DatoOpnot.Substring(0, 10) + Environment.NewLine;
                                    }
                                    else
                                    {
                                        bookSave += " \t" + DatoOpnot + Environment.NewLine;
                                    }
                                    gemSelskab = SelskabsNr;
                                }
                            }
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.Close();
            //this.Events.Dispose();
        }
    }

}


