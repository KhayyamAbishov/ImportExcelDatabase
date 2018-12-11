using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportExcelDatabase
{
    public partial class Form1 : Form
    {
       static OpenFileDialog od = new OpenFileDialog();
        string excelFilePath = od.FileName;
        public Form1()
        {

            InitializeComponent();
            

        }

        private void InsertExcelRecords()
        {
            SqlConnection sqlConnection = new SqlConnection();

            //  ExcelConn(_path);  
            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", excelFilePath);
            OleDbConnection Econ = new OleDbConnection(constr);
                string Query = string.Format("Select [StudentId] [StudentName],[StudentAge] FROM [{0}]", "Sheet1$");
                OleDbCommand Ecom = new OleDbCommand(Query, Econ);
                Econ.Open();

                DataSet ds = new DataSet();
                OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
                Econ.Close();
                oda.Fill(ds);
                DataTable Exceldt = ds.Tables[0];

                for (int i = Exceldt.Rows.Count - 1; i >= 0; i--)
                {
                    if (Exceldt.Rows[i]["Employee Name"] == DBNull.Value || Exceldt.Rows[i]["Email"] == DBNull.Value)
                    {
                        Exceldt.Rows[i].Delete();
                    }
                }
                Exceldt.AcceptChanges();
                //creating object of SqlBulkCopy      
                System.Data.SqlClient.SqlBulkCopy objbulk = new System.Data.SqlClient.SqlBulkCopy(sqlConnection);
                //assigning Destination table name      
                objbulk.DestinationTableName = "Student";
                //Mapping Table column    
                objbulk.ColumnMappings.Add("StudentId", "StudentId");
                objbulk.ColumnMappings.Add("StudentName", "StudentName");
                objbulk.ColumnMappings.Add("StudentAge", "StudentAge");
              

                //inserting Datatable Records to DataBase   
                sqlConnection.ConnectionString = "server = VSBS01; database = dbHRVeniteck; User ID = sa; Password = veniteck@2016"; //Connection Details    
            sqlConnection.Open();
                objbulk.WriteToServer(Exceldt);
            sqlConnection.Close();
                MessageBox.Show("Data has been Imported successfully.", "Imported", MessageBoxButtons.OK, MessageBoxIcon.Information);

            
         

        }
      

        private void btnChoose_Click(object sender, EventArgs e)
        {
            if (od.ShowDialog()==DialogResult.OK)
            {
                string excelFilePath = od.FileName;

            }
        }
    }
}
