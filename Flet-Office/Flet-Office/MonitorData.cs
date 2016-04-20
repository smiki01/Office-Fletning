using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Flet_Office
{
    class MonitorData
    {
        private static string sConnectionString = Properties.Settings.Default.Common_ConnectionString_Monitor;
        public void AddMonitor(string Server, string Job, string Tekst, int Alarm, string ServerGruppe)
        {
            SqlConnection objConn = new SqlConnection(sConnectionString);
            objConn.Open();

            using (SqlCommand cmd =
            new SqlCommand("INSERT INTO [pro-plz-sql01].plazait.dbo.serverlog (server,job,value,Alarm,ServerGruppe) values ('"+Server +"','"+Job +"','" +Tekst +"'," +Alarm +",'" +ServerGruppe +"')", objConn))
            {
                int rows = cmd.ExecuteNonQuery();
            }
        }
    }
}
