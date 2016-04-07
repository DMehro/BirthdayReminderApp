using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;
using System.Data;

namespace Birthday_Reminder
{
    class Program
    {
        static void Main(string[] args)
        {
            string excelPath = @"Z:\Team_Details.xlsx";
            ReadExcel(excelPath);
            
        }

        private static void ReadExcel(string excelPath)
        {
            EmailHelper emailHelper = new EmailHelper();
            OleDbConnection oledbConn =new OleDbConnection();
            try
            {
                string path = Path.GetFullPath(excelPath);
                if (Path.GetExtension(path) == ".xls")
                {
                    string tempxlsCon= string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"",path);
                    oledbConn = new OleDbConnection(tempxlsCon);
                }

                else if(Path.GetExtension(path) == ".xlsx")
                {
                    string tempxlsxPath = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=1;'", path);
                    oledbConn = new OleDbConnection(tempxlsxPath);

                }

                oledbConn.Open();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", oledbConn);
                OleDbDataAdapter oleda = new OleDbDataAdapter(cmd);
                DataSet ds = new DataSet();
                oleda.Fill(ds);

                FindUpComingBirthday(ds, emailHelper);
            }
            catch (Exception ex)
            {

            }
        }

        private static void FindUpComingBirthday(DataSet ds, EmailHelper emailHelper)
        {
            DateTime currentDate = DateTime.Now;

            for (int i = 1; i < ds.Tables[0].Rows.Count; i++)
            {
               if (Convert.ToDateTime(ds.Tables[0].Rows[i]["DOB"]).Month == DateTime.Now.Month && Convert.ToDateTime(ds.Tables[0].Rows[i]["DOB"]).Day == DateTime.Now.Date.AddDays(1).Day)
                {
                    emailHelper.SendEmail(ds.Tables[0].Rows[i]["Name"].ToString(),ds.Tables[0].Rows[i]["DOB"].ToString());
                   
                }
            }

        }
    }
}
