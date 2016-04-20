using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Flet_Office
{
    class Medlem
    {
        private string _medlemsNr;
        public string MedlemsNr
        {
            get { return _medlemsNr; }
            set { _medlemsNr = value; }
        }
        private string _datoBetOpnot;
        public string DatoBetOpnot
        {
            get { return _datoBetOpnot; }
            set { _datoBetOpnot = value; }
        }
        private string _datoOpnot;
        public string DatoOpnot
        {
            get { return _datoOpnot; }
            set { _datoOpnot = value; }
        }
        private string _listeMedlemskab;
        public string ListeMedlemskab
        {
            get { return _listeMedlemskab; }
            set { _listeMedlemskab = value; }
        }
        private string _status;
        public string Status
        {
            get { return _status; }
            set { _status = value; }
        }
        private string _lMType;
        public string LMType
        {
            get { return _lMType; }
            set { _lMType = value; }
        }

        public Medlem()
        {
            // TODO: Complete member initialization
        }
    }
}
