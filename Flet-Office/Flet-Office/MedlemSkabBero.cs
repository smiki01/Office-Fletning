using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Flet_Office
{
    class MedlemskabBero
    {
        private List<MedlemBero> alleMedl = new List<MedlemBero>();
        public List<MedlemBero> AlleMedl
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

        public MedlemskabBero()
        {
            MedlData.MedlemsDataBero(this);
        }
    }
}
