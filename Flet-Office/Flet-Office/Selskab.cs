using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Flet_Office
{
    class Selskab
    {
        private string _selnavn;
        public string SelNavn
        {
            get { return _selnavn; }
            set { _selnavn = value; }
        }
        private string _afdMail;
        public string AfdMail
        {
            get { return _afdMail; }
            set { _afdMail = value; }
        }
        private string _lokalBy;
        public string LokalBy
        {
            get { return _lokalBy; }
            set { _lokalBy = value; }
        }
        public Selskab(int Sel)
        {
            // TODO: Complete Selskab initialization
        }

    }
}
