using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GoboligAddPubDocs.GoboligDocs;

namespace Flet_Office
{
    class Gobolig
    {
        GoBDocs Journal;
        public Gobolig(string IntGobUrl, string FilePDF, string GoBib)
        {
            string GoUrl = "http://gobolig/cases" + IntGobUrl;

            Journal = new GoBDocs(GoUrl, GoBib, "", FilePDF, "", false, "", "");
            if (Journal.StatusCode != 0)
            {
                MyVars.GoJournalOk = false;
            }
        }
    }
}
