using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RASW;
using System.Data.OleDb;
using System.IO;
using System.Data.SqlClient;

namespace PricingTool
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            string fileName = "Polarcirkel Parts List for new template.xlsx";
            DataTable data = ReadSpreadsheet(Path.Combine(Directory.GetCurrentDirectory() + @"\Spreadsheets", fileName), null);
            dataGridView1.DataSource = data;

            List<string> list = ProcessDataToList(data);

            DateTime dt = new DateTime();
            double xChNOKtoGB = 0;
            CreateInsertSQL(list, DateTime.Now,xChNOKtoGB);
        }

        private void CreateInsertSQL(List<string> list, DateTime date, double ExchangeRateNOKtoGB)
        {
            string tempInsertSQL = Path.Combine(Directory.GetCurrentDirectory(), "insert.sql");

            string sql = "SET DATEFORMAT dmy; INSERT INTO[dbo].[PolarcirkleCostPrices] ([polarPriceListDate],[polarExchangeRateUsedID],[polarPartNumber],[polarAkvaPartNumber],[polarAkvaPartNames],[polarMerdedler],[polarPartDescription],[polarUnit],[polarAGcostNOK],[polarAGcostOnRequest],[polarCostGBP],[polarDateTimeAdded],[polarValid]) VALUES (";

            string sqlEnd = "')"; // SELECT SCOPE_IDENTITY();";

            if (File.Exists(tempInsertSQL)) File.Delete(tempInsertSQL);

            try
            {
                foreach (string line in list)
                {
                    if (line != "|||||||")
                    {
                        string[] split = line.Split('|');

                        if (split[0].Substring(0, 5).ToLower() != "polar")
                        {
                            string sqlData = "'" + date.ToShortDateString() + "','" + ExchangeRateNOKtoGB + "','" + split[0].ToString() + "','" + split[1].ToString() + "','" + split[2].ToString() + "','" + split[3].ToString() + "','" + split[4].ToString() + "','" + split[5].ToString() + "','" + split[6].ToString() + "','" + split[7].ToString() + sqlEnd;
                            File.AppendAllText(tempInsertSQL, sql + sqlData + Environment.NewLine);
                        }
                       
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
            }

            //try
            //{
            //    using (SqlConnection cnn = new SqlConnection(ConnectionString))
            //    {
            //        cnn.Open();
            //        using (SqlCommand cmd = new SqlCommand(scriptLine.ToString(), cnn))
            //        {
            //           int lastID = (int)cmd.ExecuteScalar();
            //           cmd.Dispose();
            //           cnn.Close();
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    //lstOutput.Items.Add("");
            //    //lstOutput.Items.Add("ERROR: " + ex.Message);
            //    //lstOutput.Items.Add("");
            //    //return "";
            //}
        }

        private List<string> ProcessDataToList(DataTable data)
        {
            List<string> sheetRows = new List<string>();
                      
            try
            {
                foreach (DataColumn col in data.Columns)
                {
                    foreach (DataRow row in data.Rows)
                    {
                        //listBox1.Items.Add(row[0] + "|" + row[1] + "|" + row[2] + "|" + row[3] + "|" + row[4] + "|" + row[5] + "|" + row[6] + "|" + row[7]);
                        sheetRows.Add(row[0] + "|" + row[1] + "|" + row[2] + "|" + row[3] + "|" + row[4] + "|" + row[5] + "|" + row[6] + "|" + row[7]);
                    }
                }

                return sheetRows;
            }
            catch (Exception)
            {
                return null;
            }
        }

        DataTable ReadSpreadsheet(string SpreadsheetFullPath, string SheetName)
        {
            //Polarcirkel Parts List for new template.xlsx

            if (SheetName == null) { SheetName = "Sheet1"; }

            DataTable dt = new DataTable();
            try
            {
                using (OleDbConnection conn = new OleDbConnection())
                {
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + SpreadsheetFullPath + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
                    using (OleDbCommand comm = new OleDbCommand())
                    {
                        comm.CommandText = "Select * from [" + SheetName + "$]";
                        comm.Connection = conn;

                        using (OleDbDataAdapter da = new OleDbDataAdapter())
                        {
                            da.SelectCommand = comm;
                            da.Fill(dt);
                        }
                    }
                    conn.Close();
                }
                return dt;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

    }
}
