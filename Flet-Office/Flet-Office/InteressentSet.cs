using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Flet_Office
{
    class InteressentSet
    {
        private List<Interessent> alleInt = new List<Interessent>();
        public List<Interessent> AlleInt
        {
            get { return alleInt; }
            set { alleInt = value; }
        }
        EGBoligData IntData = new EGBoligData();

        public InteressentSet(int MedlemNr, string KontorNavn)
        {
            IntData.InteressentData(this, MedlemNr, KontorNavn);
        }
    }
}
