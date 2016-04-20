using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Flet_Office
{
    class MedlemBero
    {
        private string _medlemsNr;
        public string MedlemsNr
        {
            get { return _medlemsNr; }
            set { _medlemsNr = value; }
        }
        private string _listeMedlemskab;
        public string ListeMedlemskab
        {
            get { return _listeMedlemskab; }
            set { _listeMedlemskab = value; }
        }

        public MedlemBero()
        {
            // TODO: Complete member initialization
        }
    }
}
