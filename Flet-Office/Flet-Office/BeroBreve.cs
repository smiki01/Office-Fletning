using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading.Tasks;
using System.EnterpriseServices;
using Microsoft.Office.Interop.Word;

namespace Flet_Office
{
    public partial class BeroBreve : Form
    {
        MedlemskabBero MedlemBero;
        Medlemskab Medlem;
        SelskabSet Selskab;
        InteressentSet Interessent;

        string PgmLog = "";
        string outputFolderPath = Properties.Settings.Default.BeroBreve_Folder;
        string bookSave = "";
        string gemMedl = "";
        string gemSelskab = "";
        string LogText = "";
        string SelNavn = "";
        string KontorNavn = "";
        int FejlCnt = 0;
        int MedlCnt = 0;
        string MailLog = Properties.Settings.Default.Common_MailLog;
        string LokalBy = "";
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
        string AfdMail = "";
        string emailModt = "";
        string IntGobUrl = "";
        string GoBib = Properties.Settings.Default.Common_GoDokBib;
        string IntNavn = "";
        string IntNr = "";
        string IntCo = "";
        string IntAdresse = "";
        string IntLokalBy = "";
        string IntPostBy = "";
        string IntInternetNr = "";
        string IntPassword = "";
        string PostStregkode = "";
        string SelskabsNr = "";
        string PgmLogSys = "";
        int IngenMailBogm = 0;
        string FilePDF = ""; 

        public BeroBreve()
        {
            InitializeComponent();
            if (MyVars.ParmYes == "AUTO")
            {
                this.timer1.Enabled = true;
                this.timer1.Interval = 5000;
                this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            }
            MedlemBero = new MedlemskabBero();
            string skabelonPath = Properties.Settings.Default.Common_SkabelonSti;
            string skabelonOriginal = Properties.Settings.Default.BeroBreve_SkabelonNavn;

            string outputFolderOld = Properties.Settings.Default.BeroBreve_FolderOld;
            string dt = DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss");
            PgmLog = Properties.Settings.Default.BeroBreve_ErrorLog + dt + ".Log";
            PgmLogSys = "C:\\FletOffice\\ProgramError_Opskriv_SysLog_" + dt + ".Log";

            if (Directory.Exists(outputFolderOld))
            {
                System.IO.Directory.Delete(outputFolderOld, true);
            }
            System.IO.Directory.Move(outputFolderPath, outputFolderOld);
            System.IO.Directory.CreateDirectory(outputFolderPath);

            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();

            foreach (MedlemBero oMedlemsDataBero in MedlemBero.AlleMedl)      //her løbes alle dagens nye bero medlemmer igennem
            {
                int cnt = 0;
                Medlem = new Medlemskab(oMedlemsDataBero.MedlemsNr);
                foreach (Medlem oMedlemsData in Medlem.AlleMedl)      //her udtrækkes medlemskaber pr. medlem
                {
                    //her behandles medlemsdata
                    Document skabelon = application.Documents.Open(skabelonPath + skabelonOriginal);
                    application.Visible = false;

                    if (gemMedl != oMedlemsData.MedlemsNr && bookSave != "" && gemMedl != "")
                    {
                        //Her behandles multiline bogmærker, PDF generering, Email afsendelse og journalisering til Gobolig

                        BehandlMedlem(skabelon, gemMedl, oMedlemsData.LMType);

                        Document document = application.Documents.Open(skabelonPath + skabelonOriginal);
                        application.Visible = false;

                        // clone the doc object.
                        skabelon = document;
                    }
                    else
                    {
                        if (gemMedl != oMedlemsData.MedlemsNr && gemMedl != MyVars.Selopdat && gemMedl != "")
                        {
                            textBox1.Text += "Nu har vi behandlet medlem nr.: " + gemMedl + ", som kun er medlem af selskab 99 og derfor ikke skal have et berobrev tilsendt." + Environment.NewLine;
                            LogText = "Nu har vi behandlet medlem nr.: " + gemMedl + ", som kun er medlem af selskab 99 og derfor ikke skal have et berobrev tilsendt.";
                            File.AppendAllText(PgmLog, LogText + Environment.NewLine);
                            MyVars.Selopdat = "";
                            // her skal Word dokument slettes og hoppes videre til næste medlem.
                            ((_Document)skabelon).Close(WdSaveOptions.wdDoNotSaveChanges);   //her lukkes skabelonen

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
                    OpdaterBookmark(skabelon, oMedlemsData.MedlemsNr, oMedlemsData.ListeMedlemskab, oMedlemsData.Status, oMedlemsData.LMType);
                    cnt++;

                    gemMedl = oMedlemsData.MedlemsNr;
                    if (cnt == Medlem.AntalMedl)           // her behandles sidste medlem i 2. SQL udtræk færdigt
                    {
                        //Her behandles multiline bogmærker, PDF generering, Email afsendelse, journalisering til Gobolig og lukning af skabelon.
                        BehandlMedlem(skabelon, gemMedl, oMedlemsData.LMType);
                    }
                    MyVars.InsertOk = false;
                    gemMedl = "";
                    gemSelskab = "";
                }
            }
            //Her afsluttes programmet.....
            string Server = Properties.Settings.Default.Common_MonitorServerNavn;
            string Job = Properties.Settings.Default.BeroBreve_MonitorJobNavn;
            string ServerGruppe = Properties.Settings.Default.Common_MonitorGruppeNavnGO;
            string Tekst = "";
            int Alarm = Properties.Settings.Default.Common_MonitorAlarmVærdiOK;
            int antMailSendt = MedlCnt - FejlCnt;

            if (File.Exists(PgmLogSys))
            {
                MyVars.emailSubj = "Systemfejllog for Berobreve";
                MyVars.emailBody = "Dette er en fejllog som viser system relaterede fejl ved kørsel af Berobreve. " + Environment.NewLine + "Se vedhæftede log for detaljer";
                SendEmail EmailObj = new SendEmail("", "", "admin@bdk.dk", PgmLogSys, "sa-gobolig@bdk.dk", MailLog, "");
            }
            if (File.Exists(PgmLog))
            {
                string afs = Properties.Settings.Default.Common_EmailAfsFejl;
                string modt = Properties.Settings.Default.Common_EmailEGModtLog;

                if (FejlCnt > 0)
                {
                    MyVars.emailSubj = Properties.Settings.Default.BeroBreve_EmailSubjFejl;
                    MyVars.emailBody = Properties.Settings.Default.Common_EmailBodyFejl + " (BeroBreve) " + Environment.NewLine + "Se vedhæftede log for detaljer";
                    Tekst = Properties.Settings.Default.Common_MonitorTekstFejl + FejlCnt + ". " + Properties.Settings.Default.Common_MonitorTekst + antMailSendt;
                    Alarm = Properties.Settings.Default.Common_MonitorAlarmVærdiFejl;
                }
                else
                {
                    MyVars.emailBody = Properties.Settings.Default.Common_EmailBodyOk + " (BeroBreve) " + Environment.NewLine + "Se vedhæftede log for detaljer";
                    MyVars.emailSubj = "Berobreve: " + Properties.Settings.Default.Common_EmailSubjOk;
                    Tekst = Properties.Settings.Default.Common_MonitorTekstOK + ". " + Properties.Settings.Default.Common_MonitorTekst + antMailSendt;
                }
                LogText = Tekst;
                File.AppendAllText(PgmLog, LogText + Environment.NewLine);

                SendEmail EmailObj = new SendEmail("", "", afs, PgmLog, modt, MailLog, "");
            }
            else
            {
                Tekst = Properties.Settings.Default.Common_MonitorTekstFejl + FejlCnt + ". " + Properties.Settings.Default.Common_MonitorTekst + antMailSendt;
            }

            if (MedlCnt == 0)
            {
                textBox1.Text += Properties.Settings.Default.BeroBreve_SlutTekstIngen + Environment.NewLine;
                Tekst = Properties.Settings.Default.BeroBreve_SlutTekstIngen + Environment.NewLine;
            }

            MonitorAdd MonitorAddGO = new MonitorAdd(Server, Job, Tekst, Alarm, ServerGruppe);
            ServerGruppe = Properties.Settings.Default.Common_MonitorGruppeNavnEG;
            MonitorAdd MonitorAddEG = new MonitorAdd(Server, Job, Tekst, Alarm, ServerGruppe);

            // Close word.
            ((_Application)application).Quit();                                  //her lukkes Word applicationen
        }



        //Herunder følger alle funktionerne som kaldes fra hovedprogrammet
        public void BehandlMedlem(Document skabelon, string MedlemsNr, string LMType)
        {
            //her opdateres multiline bogmærkerne 
            SkrivBookmark(skabelon, bookSave, MedlemsNr, LMType);

            textBox1.Text += Properties.Settings.Default.Common_SlutTekst + gemMedl + ". (INT-" + IntNr.PadLeft(8, '0') + ")" + Environment.NewLine;

            bookSave = "";
            //her konverteres dokumentet til PDF
            string FilePDF = @outputFolderPath + "\\BeroBrev_MedlemNr_" + gemMedl + ".pdf";
            PDF PDFObj = new PDF(skabelon, FilePDF);

            //her vedhæftes det udfyldte PDF dokument til en email og derefter slettes det og mail sendes, herefter genoptages loop (break)
            UdtrKolOpl(gemMedl, LokalBy);          //her hentes afdelingskontors email adresse (til afsender)
            if (AfdKontorEmail == "" || AfdKontorEmail.IndexOf("Afdelingskontoret") == -1)
            {
                AfdKontorEmail = AfdMail;
            }
            string Kategori = "0002009999";

            SendEmail EmailObj = new SendEmail(gemMedl, IntNavn, AfdKontorEmail, FilePDF, emailModt, MailLog, Kategori);
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
                    MyVars.emailBody += Environment.NewLine + "Send bero-brev med posten i stedet for til:" + Environment.NewLine +
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

                    IngenMailBogm = 0;
                }
            }
            else
            {
                int count = MedlCnt + 1;
                LogText = count + ". Bero-brev sendt ok på e-mail til medlem: " + IntNavn + " (99-" + gemMedl + ") på denne e-mail adresse: " + emailModt + Environment.NewLine;
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
            ((_Document)skabelon).Close(WdSaveOptions.wdDoNotSaveChanges);   //her lukkes skabelonen
            MedlCnt++;
            gemMedl = "";
            gemSelskab = "";
            bookSave = "";
        }


        public void SkrivBookmark(Document skabelon, string bookSave, string MedlemsNr, string LMType)
        {
            foreach (Bookmark bookmark in skabelon.Bookmarks)
            {
                //her isættes funden tekst for flere liniers bogmærker
                string bMark = bookmark.Name;
                switch (bMark)
                {
                    case "Medlem_Liste_MedlemSelskab":
                        bookmark.Range.Font.Size = 9.50F;
                        bookmark.Range.Bold = 2;
                        bookmark.Range.Text = bookSave;
                        MyVars.Selopdat = MedlemsNr;
                        break;
                    case "Medlem_Liste_NLMType":
                        bookmark.Range.Font.Size = 8.00F;
                        bookmark.Range.Bold = 2;
                        bookmark.Range.Text = "";
                        break;
                    case "Medlem_Liste_Status":
                        bookmark.Range.Font.Size = 8.00F;
                        bookmark.Range.Bold = 2;
                        bookmark.Range.Text = "";
                        break;
                    case "Ingenmail_Tekst":
                        if (emailModt != "")
                        {
                            bookmark.Range.Text = "Bemærk at hvis du modtager dette dokument som brev, er det fordi vi ikke har kunne sende det på den e-mail adresse vi har fået fra dig: " + emailModt;
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
        //emailModt = "smiki01@hotmail.dk";          //kun til test brug - invalid e-mail adresse
        //emailModt = "";                            //kun til test brug - manglende e-mail adresse
                PostStregkode = oInteressentData.PostStregkode;
            }
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

        public void OpdaterBookmark(Document skabelon, string MedlemsNr, string ListeMedlemskab, string Status, string LMType)
        {
            foreach (Bookmark bookmark in skabelon.Bookmarks)
            {
                //her isættes funden tekst for bogmærker pr. medlem
                string bMark = bookmark.Name;
                switch (bMark)
                {
                    case "Ingenmail_Tekst":
                        if (IngenMailBogm < 1)
                        {
                            if (emailModt == "")
                            {
                                bookmark.Range.Text = Properties.Settings.Default.Common_EmailModtIngen1 + Environment.NewLine +
                                Properties.Settings.Default.Common_EmailModtIngen2 + Environment.NewLine;
                            }
                            IngenMailBogm++;
                        }
                        break;
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
                        bookmark.Range.Text = MedlemsNr;
                        break;
                    case "Firma_Navn":
                        bookmark.Range.Text = FirmaNavn;
                        break;
                    case "Medlem_Navn":
                        bookmark.Range.Text = IntNavn;
                        break;
                    case "Medlem_CO":
                        bookmark.Range.Text = IntCo;
                        break;
                    case "Medlem_Adresse":
                        bookmark.Range.Text = IntAdresse;
                        break;
                    case "Medlem_Lokalby":
                        bookmark.Range.Text = IntLokalBy;
                        break;
                    case "Medlem_Postby":
                        bookmark.Range.Text = IntPostBy;
                        break;
                    case "Medlem_Internetnr":
                        if (IntInternetNr == "")
                        {
                            IntInternetNr = IntNr;
                        }
                        bookmark.Range.Font.Size = 9.50F;
                        bookmark.Range.Bold = 2;
                        bookmark.Range.Text = IntInternetNr;
                        break;
                    case "Dato_side1":
                        bookmark.Range.Font.Size = 9.50F;
                        bookmark.Range.Text = DateTime.Today.ToString("dd. MMMM yyyy");
                        break;
                    case "Post_Stregkode":
                        bookmark.Range.Font.Size = 9.50F;
                        bookmark.Range.Bold = 2;
                        bookmark.Range.Text = PostStregkode;
                        break;
                    case "Medlem_Liste_MedlemSelskab":
                        if (ListeMedlemskab != "99" && gemSelskab != SelskabsNr && SelNavn != "")
                        {
                            bookmark.Range.Font.Size = 9.50F;
                            bookmark.Range.Bold = 2;
                            if ((gemMedl == MedlemsNr && (Status == "0" || Status == "1")) || gemMedl == "")
                            {
                                if (bookmark.Name == "Medlem_Liste_MedlemSelskab")
                                {
                                    bookSave += " " + SelNavn + ": ";
                                }
                            }
                        }
                        break;
                    case "Medlem_Liste_NLMType":
                        if (ListeMedlemskab != "99" && gemSelskab != SelskabsNr && SelNavn != "")
                        {
                            bookmark.Range.Font.Size = 8.00F;
                            bookmark.Range.Bold = 2;
                            if ((gemMedl == MedlemsNr && (Status == "0" || Status == "1")) || gemMedl == "")
                            {
                                if (bookmark.Name == "Medlem_Liste_NLMType")
                                {
                                    string LM = "";
                                    string bLM = LMType;
                                    switch (bLM)
                                    {
                                        case "1":
                                            LM = "Familie ";
                                            break;
                                        case "4":
                                            LM = "Ungdom  ";
                                            break;
                                        case "5":
                                            LM = "Ældre   ";
                                            break;
                                        case "6":
                                            LM = "Pleje   ";
                                            break;
                                        case "7":
                                            LM = "1.vær.  ";
                                            break;
                                        case "8":
                                            LM = "Ungdom  ";
                                            break;
                                        case "10":
                                            LM = "Andel   ";
                                            break;
                                        case "12":
                                            LM = "Klubvær. ";
                                            break;
                                        default:
                                            LM = "Ukendt   ";
                                            break;
                                    }
                                    bookSave += " \t" + LM;
                                }
                            }
                        }
                        break;
                    case "Medlem_Liste_Status":
                        if (ListeMedlemskab != "99" && gemSelskab != SelskabsNr && SelNavn != "")
                        {
                            bookmark.Range.Font.Size = 8.00F;
                            bookmark.Range.Bold = 2;
                            if ((gemMedl == MedlemsNr && (Status == "0" || Status == "1")) || gemMedl == "")
                            {
                                if (bookmark.Name == "Medlem_Liste_Status")
                                {
                                    string Stat = "";
                                    string bStat = Status;
                                    switch (bStat)
                                    {
                                        case "0":
                                            Stat = "Aktiv  ";
                                            break;
                                        case "1":
                                            Stat = "Bero   ";
                                            break;
                                        case "8":
                                            Stat = "Afv.bet.";
                                            break;
                                        case "9":
                                            Stat = "Afv.bet.";
                                            break;
                                        default:
                                            Stat = "Ukendt  ";
                                            break;
                                    }
                                    bookSave += " \t" + Stat + Environment.NewLine; ;
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
