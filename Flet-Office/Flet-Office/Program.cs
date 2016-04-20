using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Flet_Office
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Count() > 0)
            {
                if (args.Count() == 2)
                {
                    MyVars.ParmYes = args[1];                    //Her gemmes 2. parameter (AUTO), hvis den er udfyldt
                }
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                switch (args[0])
                {
                    case "/?":
                        MessageBox.Show("Parametre til Flet-Office programmet:" + Environment.NewLine + Environment.NewLine + 
                        "1. parameter kan være 'Opskrivningsbevis', 'BeroBreve' eller '/?'" + Environment.NewLine + 
                        "2. parameter kan være 'AUTO'. Betyder at programmet køres uden at Windowsform vises." + Environment.NewLine + Environment.NewLine +
                        "Bemærk at Flet-Office SKAL kaldes med mindst 1 af ovennævnte parametre!");
                        break;
                    case "Opskrivningsbevis":
                        Application.Run(new Opskrivningsbevis());
                        break;
                    case "BeroBreve":
                        Application.Run(new BeroBreve());
                        break;
                    default:
                        MessageBox.Show("Flet-Office kaldt med forkerte parametre=(" + args[0] + "). Start Flet-Office med '/?' for at se korrekte parametre.");
                        break;
                }                
            }
            else
            { MessageBox.Show("Flet-Office kaldt uden parametre. Start Flet-Office med '/?' for at se korrekte parametre."); }
        }
    }
    public static class MyVars
    {
        public static bool InsertOk = false;
        public static bool GoJournalOk = true;
        public static string ParmYes = "";
        public static bool EmailSendtOK = false;
        public static string emailSubj = "";
        public static string emailBody = "";
        public static string Selopdat = "";
        public static string MailFejltekst = "";
    }
}
