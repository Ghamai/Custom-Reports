using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace Custom_Reports
{
    class Connection
    {
        public string shipingnum  { get; set;}
        public string poamount { get; set;}

        public DataTable Tble()
        {
            string url;
            url = "O:/TASB Shared/BuyBoard/BuyBoard vendor files/Copy of Vendor list.xls";
            //url = "C:/Users/khuragha/Desktop/Copy of Vendor list.xls";

            string pathconn = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" + url + ";Extended Properties =\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection connect = new OleDbConnection(pathconn);
            OleDbDataAdapter datadap = new OleDbDataAdapter("Select*from[Member Upload$]", connect);
            DataTable dt = new DataTable();
            datadap.Fill(dt);
            connect.Close();
            return dt;

        }

        public DataTable TbleVenReport()
        {
            string url;
            url = "O:/TASB Shared/BuyBoard/BuyBoard vendor files/Groups.xls";
            //url = "C:/Users/khuragha/Documents/Groups.xls";

            string pathconn = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" + url + ";Extended Properties =\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection connect = new OleDbConnection(pathconn);
            //OleDbDataAdapter datadap = new OleDbDataAdapter("Select*from[Member Upload$]", connect);
            OleDbDataAdapter datadap = new OleDbDataAdapter("Select*from[Source Records$]", connect);
            DataTable dt = new DataTable();
            datadap.Fill(dt);
            connect.Close();
            return dt;

        }


    }
}
