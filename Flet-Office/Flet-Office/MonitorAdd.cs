using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Flet_Office
{
    class MonitorAdd
    {
        MonitorData EGCon = new MonitorData();

        public MonitorAdd(string Server, string Job, string Tekst, int Alarm, string ServerGruppe)
        {
            EGCon.AddMonitor(Server, Job, Tekst, Alarm, ServerGruppe);
        }
    }
}
