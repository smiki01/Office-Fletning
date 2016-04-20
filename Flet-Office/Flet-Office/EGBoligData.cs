using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Flet_Office
{
    class EGBoligData
    {
        private string _exception;
        public string Exception
        {
            get { return _exception; }
            set { _exception = value; }
        }

        private static string sConnectionStringSQL02 = Properties.Settings.Default.Common_ConnectionString_SQL;

        public void MedlemsDataBero(MedlemskabBero oMedlemskab)
        {
            DataTable tblMedlem;
            SqlConnection objConn = new SqlConnection(sConnectionStringSQL02);
            objConn.Open();
            SqlDataAdapter daMedlem = new SqlDataAdapter(Properties.Settings.Default.BeroBreve_Select, objConn);
            //SqlDataAdapter daMedlem = new SqlDataAdapter("SELECT  DISTINCT TOP (100) PERCENT Medlem FROM dbo.MedlemAfSelskabStatusLog WHERE MedlemSelskab <> 99 AND (medlem = 108444 or medlem = 101937 or medlem = 112879) ORDER BY Medlem", objConn);
            DataSet dsMedlem = new DataSet("MedlemTab");
            daMedlem.FillSchema(dsMedlem, SchemaType.Source, "MedlemAfSelskabStatusLog");
            daMedlem.Fill(dsMedlem, "MedlemAfSelskabStatusLog");
            tblMedlem = dsMedlem.Tables["MedlemAfSelskabStatusLog"];
    
            objConn.Close();
       
            oMedlemskab.AntalMedl = tblMedlem.Rows.Count;

            foreach (DataRow drMedlem in tblMedlem.Rows)
            {
                try
                {
                    MedlemBero oMedlem = new MedlemBero();
                    oMedlem.MedlemsNr = drMedlem["Medlem"].ToString();

                    oMedlemskab.AlleMedl.Add(oMedlem);

                }
                catch (Exception ex)
                {
                    throw new Exception("Der er opstået en fejl, i forbindelse med hentning af data fra MedlemAfSelskabStatusLog tabellen i Bolig databasen", ex);
                }
            }
        }
        public void MedlemsData(Medlemskab oMedlemskab, string MedlemsNr)
        {
            DataTable tblMedlem;
            SqlConnection objConn = new SqlConnection(sConnectionStringSQL02);
            objConn.Open();
            //Hvis der kaldes fra Bero breve, trækkes der kun ud på et enkelt medlemsnr.
            SqlDataAdapter daMedlem = new SqlDataAdapter("SELECT * FROM [Bolig].[dbo].[MedlemAfSelskab] where medlem = '" + MedlemsNr + "' and Status <> 3 order by Medlem, medlemselskab", objConn);  

            if(MedlemsNr == "")
            {
                //Hvis der kaldes fra Opskrivningsbevis, trækkes ud på alle medlemsnr. hvor opnoteringsdatoen er fra igår 
                daMedlem = new SqlDataAdapter(Properties.Settings.Default.Opskrivningsbevis_Select, objConn);
                //daMedlem = new SqlDataAdapter("SELECT DISTINCT dbo.MedlemAfSelskab.Medlem, MedlemAfSelskab_1.Sel, MedlemAfSelskab_1.Medlem AS Expr1, MedlemAfSelskab_1.MedlemSelskab, MedlemAfSelskab_1.Status, MedlemAfSelskab_1.DatoOpnot, MedlemAfSelskab_1.DatoBetOpnot, MedlemAfSelskab_1.Lmtype FROM dbo.MedlemAfSelskab RIGHT OUTER JOIN dbo.MedlemAfSelskab AS MedlemAfSelskab_1 ON dbo.MedlemAfSelskab.Medlem = MedlemAfSelskab_1.Medlem where datediff(day,MedlemAfSelskab_1.DatoOpnot,getdate()) = 3 and MedlemAfSelskab_1.Status <> 3 and MedlemAfSelskab_1.DatoBetOpnot is not null and (MedlemAfSelskab_1.Medlem = 99753) and datediff(day,MedlemAfSelskab_1.DatoBetOpnot,getdate()) = 3 ORDER BY Medlem, MedlemSelskab, DatoOpnot desc", objConn);
                //  and (MedlemAfSelskab_1.Medlem = 22563 or MedlemAfSelskab_1.Medlem = 59261 or MedlemAfSelskab_1.Medlem = 77924) and datediff(day,MedlemAfSelskab_1.DatoBetOpnot,getdate()) = 7  
            }

            DataSet dsMedlem = new DataSet("MedlemTab");
            daMedlem.FillSchema(dsMedlem, SchemaType.Source, "MedlemAfSelskab");
            daMedlem.Fill(dsMedlem, "MedlemAfSelskab");
            tblMedlem = dsMedlem.Tables["MedlemAfSelskab"];

            objConn.Close();

            oMedlemskab.AntalMedl = tblMedlem.Rows.Count;

            foreach (DataRow drMedlem in tblMedlem.Rows)
            {
                try
                {
                    Medlem oMedlem = new Medlem();
                    oMedlem.MedlemsNr = drMedlem["Medlem"].ToString();
                    oMedlem.DatoBetOpnot = drMedlem["DatoBetOpnot"].ToString();
                    oMedlem.DatoOpnot = drMedlem["DatoOpnot"].ToString();
                    oMedlem.ListeMedlemskab = drMedlem["MedlemSelskab"].ToString();
                    oMedlem.Status = drMedlem["Status"].ToString();
                    oMedlem.LMType = drMedlem["Lmtype"].ToString();
                    oMedlemskab.AlleMedl.Add(oMedlem);

                }
                catch (Exception ex)
                {
                    throw new Exception("Der er opstået en fejl, i forbindelse med hentning af data fra MedlemAfSelskab tabellen i Bolig databasen", ex);
                }
            }
        }
        public void SelskabsData(SelskabSet oSelskabSet, int Sel)
        {
            SqlConnection objConn = new SqlConnection(sConnectionStringSQL02);
            objConn.Open();

            SqlDataAdapter daSelskab = new SqlDataAdapter("SELECT * FROM [Bolig].[dbo].[Selskab] where sel = " + Sel, objConn);
            DataSet dsSelskab = new DataSet("SelskabTab");
            daSelskab.FillSchema(dsSelskab, SchemaType.Source, "Selskab");
            daSelskab.Fill(dsSelskab, "Selskab");
            objConn.Close();

            DataTable tblSelskab;
            tblSelskab = dsSelskab.Tables["Selskab"];

            foreach (DataRow drSelskab in tblSelskab.Rows)
            {
                try
                {
                    Selskab oSelskab = new Selskab(Sel);
                    oSelskab.SelNavn = drSelskab["Navn"].ToString();
                    oSelskab.AfdMail = drSelskab["Email"].ToString();
                    oSelskab.LokalBy = drSelskab["Lokalby"].ToString();
                    //MessageBox.Show("Data hentes: " + oMedlem.ListeMedlemskab);
                    oSelskabSet.AlleSel.Add(oSelskab);

                }
                catch (Exception ex)
                {
                    throw new Exception("Der er opstået en fejl, i forbindelse med hentning af data fra Selskab tabellen i Bolig databasen", ex);
                }
            }
        }
        public void InteressentData(InteressentSet oInteressentSet, int MedlemNr, string KontorNavn)
        {
            SqlConnection objConn = new SqlConnection(sConnectionStringSQL02);
            objConn.Open();
            DataSet dsInteressent = new DataSet("InteressentTab");
            if(KontorNavn == "")
            {
                SqlDataAdapter daInteressent = new SqlDataAdapter("SELECT * FROM [Bolig].[dbo].[Interessentadresse] where Interessentnr = (SELECT top 1 Interessentnr from [Bolig].[dbo].[Medlem] where type = 'I' and medlem = " + MedlemNr + ")", objConn);
                daInteressent.FillSchema(dsInteressent, SchemaType.Source, "Interessent");
                daInteressent.Fill(dsInteressent, "Interessent");
            }
            else
            {
                SqlDataAdapter daInteressent = new SqlDataAdapter("SELECT * FROM [Bolig].[dbo].[Interessentadresse] where Navn = '" + KontorNavn + "'", objConn);
                daInteressent.FillSchema(dsInteressent, SchemaType.Source, "Interessent");
                daInteressent.Fill(dsInteressent, "Interessent");
            }

            objConn.Close();
            DataTable tblInteressent;
            tblInteressent = dsInteressent.Tables["Interessent"];
            objConn.Close();

            foreach (DataRow drInteressent in tblInteressent.Rows)
            {
                try
                {
                    Interessent oInteressent = new Interessent(MedlemNr);
                    oInteressent.InteressentNr = drInteressent["Interessentnr"].ToString();
                    oInteressent.IntNavn = drInteressent["Navn"].ToString();
                    oInteressent.IntCO = drInteressent["Co"].ToString();
                    oInteressent.IntAdresse = drInteressent["Adresse"].ToString();
                    oInteressent.IntLokalBy = drInteressent["Lokalby"].ToString();
                    oInteressent.IntPostBy = drInteressent["Postby"].ToString();
                    oInteressent.IntInternetNr = drInteressent["Internetnr"].ToString();
                    oInteressent.IntPassword = drInteressent["Password"].ToString();
                    oInteressent.IntGobUrl = drInteressent["gobUrl"].ToString();
                    oInteressent.IntTlf = drInteressent["Telefon_arbejde"].ToString();
                    oInteressent.IntFax = drInteressent["Telefon_Fax"].ToString();
                    oInteressent.IntEmail = drInteressent["Email1"].ToString();
                    oInteressent.IntTraftid1 = drInteressent["Træffetid1"].ToString();
                    oInteressent.IntTraftid2 = drInteressent["Træffetid2"].ToString();
                    oInteressent.IntTraftid3 = drInteressent["Træffetid3"].ToString();
                    oInteressent.IntTraftid4 = drInteressent["Træffetid4"].ToString();
                    oInteressent.IntTraftid5 = drInteressent["Træffetid5"].ToString();
                    oInteressent.PostStregkode = drInteressent["PostDanmarkStregkode"].ToString();
                    oInteressentSet.AlleInt.Add(oInteressent);
                }
                catch (Exception ex)
                {
                    throw new Exception("Der er opstået en fejl, i forbindelse med hentning af data fra Selskab tabellen i Bolig databasen", ex);
                }
            }
        }
        public Mail GetMailData(string Overskrift)
        {
            SqlConnection objConn = new SqlConnection(sConnectionStringSQL02);
            objConn.Open();

            SqlDataAdapter daMail = new SqlDataAdapter("SELECT * FROM [Bolig].[dbo].[MailTekster] where overskrift = '" + Overskrift+"'", objConn);
            DataSet dsMail = new DataSet("MailTeksterTab");
            daMail.FillSchema(dsMail, SchemaType.Source, "MailTekster");
            daMail.Fill(dsMail, "MailTekster");
            objConn.Close();
            DataTable tblMail;
            tblMail = dsMail.Tables["MailTekster"];

            Mail oMail = new Mail(Overskrift);
            if (tblMail.Rows.Count > 0)
            {
                try
                {
                    oMail.Tekst = tblMail.Rows[0]["Tekst"].ToString();
                    oMail.Emne = tblMail.Rows[0]["Emne"].ToString();
                }
                catch (Exception ex)
                {
                    throw new Exception("Der opstår en fejl, i forbindelse med at der hentes data fra MailTekster tabellen i Bolig db:" + Overskrift, ex);
                }
            }
            return oMail;
        }
        public void AddESDHMJFejl(int IntNummer, string PersId)
        {
            SqlConnection objConn = new SqlConnection(sConnectionStringSQL02);
            objConn.Open();

            //Her opdateres Interessent felt med korrekt interessentnr
            PersId = "'" + PersId + "'";
            using (SqlCommand cmd =
            new SqlCommand("INSERT INTO ESDHMJFejl (DokumentType, Interessent, Egdi, PersonID, Oprettet, KoeNr) VALUES ('Sag', '" + IntNummer + "' , '0', " + PersId + " , GETDATE() , 1)", objConn))
            {
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    MyVars.InsertOk = true;
                }
            }
        }
    }
}
