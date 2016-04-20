using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Flet_Office
{
    class SelskabSet
    {
        private List<Selskab> alleSel = new List<Selskab>();
        public List<Selskab> AlleSel
        {
            get { return alleSel; }
            set { alleSel = value; }
        }
        EGBoligData SelData = new EGBoligData();

        public SelskabSet(int Sel)
        {
            SelData.SelskabsData(this, Sel);
        }
    }
}
