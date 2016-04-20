using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Flet_Office
{
    class Medlemskab
    {
        private List<Medlem> alleMedl = new List<Medlem>();
        public List<Medlem> AlleMedl
        {
            get { return alleMedl; }
            set { alleMedl = value; }
        }
        private int _antalMedl;
        public int AntalMedl
        {
            get { return _antalMedl; }
            set { _antalMedl = value; }
        }
        EGBoligData MedlData = new EGBoligData();

        public Medlemskab(string MedlemsNr)
        {
            MedlData.MedlemsData(this, MedlemsNr);
        }
    }
}
