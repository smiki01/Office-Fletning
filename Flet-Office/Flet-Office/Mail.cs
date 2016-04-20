using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Flet_Office
{
    class Mail
    {
        private string _tekst;
        public string Tekst
        {
            get { return _tekst; }
            set { _tekst = value; }
        }
        private string _emne;
        public string Emne
        {
            get { return _emne; }
            set { _emne = value; }
        }
        public Mail(string Overskrift)           
        {
            // TODO: Complete Selskab initialization
        }
    }
}
