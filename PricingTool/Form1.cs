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
            //CreateInsertSQL(list, DateTime.Now,xChNOKtoGB);

            //List<string> plData = CreateInsertSQL(list, DateTime.Now, xChNOKtoGB);
            bool result = ExecuteSQL(CreateInsertSQL(list, DateTime.Now, xChNOKtoGB));
        }

        private bool ExecuteSQL(List<string> inserts)
        {
            string ConnectionString = "Server=tcp:akvasql1.database.windows.net,1433;Initial Catalog=PricingToolDev1;Persist Security Info=False;User ID=rwilson;Password=pCj7uu573; MultipleActiveResultSets =False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";
            try
            {
                using (SqlConnection cnn = new SqlConnection(ConnectionString))
                {
                    cnn.Open();

                    foreach (var sql in inserts)
                    {
                        try
                        {
                            using (SqlCommand cmd = new SqlCommand(sql, cnn))
                            {
                                //lastID = (int)cmd.ExecuteScalar();
                                cmd.ExecuteScalar();
                                cmd.Dispose();
                            }
                        }
                        catch (Exception exc)
                        {
                            string irwtfpn = exc.Message;
                        }
                    }
                    cnn.Close();
                }

                return true;
            }
            catch (Exception ex)
            {
                string fm = ex.Message;
               return false;
            }
        }

        private List<string> CreateInsertSQL(List<string> list, DateTime date, double ExchangeRateNOKtoGB)
        {
            List<string> SQL_Inserts = new List<string>();
            string tempInsertSQL = Path.Combine(Directory.GetCurrentDirectory(), "insert.sql");

            string sql = "SET DATEFORMAT dmy; INSERT INTO[dbo].[PolarcirkleCostPrices] ([polarPriceListDate],[polarExchangeRateUsedID],[polarPartNumber],[polarAkvaPartNumber],[polarAkvaPartNames],[polarMerdedler],[polarPartDescription],[polarUnit],[polarAGcostNOK],[polarAGcostOnRequest],[polarCostGBP],[polarDateTimeAdded],[polarValid]) VALUES (";

            if (File.Exists(tempInsertSQL)) File.Delete(tempInsertSQL);

            string polarAGcostOnRequest = "0";
            string polarCostGBP = "0.00";
            string polarAGcostNOK = "0";

            try
            {
                foreach (string line in list)
                {
                    if (line != "|||||||")
                    {
                        string[] split = line.Split('|');

                        if(split[6] == "on request")
                        {
                           polarAGcostOnRequest = "1";
                           polarCostGBP = "0.00";
                        }
                        else
                        {
                            if(split[6].ToLower().Contains("kr"))
                                polarAGcostNOK = split[6].Replace("kr", "").Trim();

                            if (polarAGcostNOK.Contains(" "))
                                polarAGcostNOK = polarAGcostNOK.Replace(" ", "");

                            if (polarAGcostNOK.Contains(","))
                                polarAGcostNOK = polarAGcostNOK.Replace(",", "");
                        }
                        
                        if(split[7].Length == 0) { polarCostGBP = "0"; }    

                        string sqlData = "'" + date.ToShortDateString() + "'," + ExchangeRateNOKtoGB + ",'" + split[0].ToString() + "','" + split[1].ToString() + "','" + split[2].ToString() + "','" + split[3].ToString() + "','" + split[4].ToString() + "','" + split[5].ToString() + "','" + polarAGcostNOK + "','" + polarAGcostOnRequest + "'," + polarCostGBP + ",'" + date.ToShortDateString() +"'," + "'1'" + ")";

                        //File.AppendAllText(tempInsertSQL, sql + sqlData + Environment.NewLine);  // debugging
                        SQL_Inserts.Add(sql + sqlData);
                    }
                }

                SQL_Inserts.RemoveAt(0);    // remove titles 
                return SQL_Inserts;
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                return null;
            }
        }

        private List<string> ProcessDataToList(DataTable data)
        {
            List<string> sheetRows = new List<string>();
                      
            try
            {
                foreach (DataRow row in data.Rows)
                {
                    //listBox1.Items.Add(row[0] + "|" + row[1] + "|" + row[2] + "|" + row[3] + "|" + row[4] + "|" + row[5] + "|" + row[6] + "|" + row[7]);
                    sheetRows.Add(row[0] + "|" + row[1] + "|" + row[2] + "|" + row[3] + "|" + row[4] + "|" + row[5] + "|" + row[6] + "|" + row[7]);
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
