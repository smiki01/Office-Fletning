using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Flet_Office
{
    class ESDHMJFejlAdd
    {
        EGBoligData EGCon = new EGBoligData();

        public ESDHMJFejlAdd(int IntNummer, string PersId)
        {
            EGCon.AddESDHMJFejl(IntNummer, PersId);
        }
    }
}
