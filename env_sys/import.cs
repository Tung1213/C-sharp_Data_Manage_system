using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace env_sys
{
    public partial class import : Form
    {
        public import()
        {
            InitializeComponent();
        }

        
        
        SqlConnection mySQLConnection = new SqlConnection("User ID=user;password=1234;Initial Catalog=Env_Data;Data Source=192.168.43.247");
     
        
        string filename=null;
        private void import_Load(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel 97-2003 (*.xls)|*.xls|文字檔 (Tab 字元分隔) (*.txt)|*.txt";

            if (ofd.ShowDialog() == DialogResult.OK)
            {

                string excelpath = ofd.FileName;
                filename = Path.GetFileNameWithoutExtension(excelpath);
                dataGridView1.DataSource = ReadFromExcel(excelpath);
            }
        }
        private DataTable ReadFromExcel(string excelpath)
        {
            string sExt = System.IO.Path.GetExtension(excelpath);
            string sConn = null;
            if (sExt == ".xlsx")
            {
                sConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + excelpath + ";" + "Extended Properties='Excel 12.0;HDR=YES'";
            }
            else if (sExt == ".xls")
            {
                sConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelpath + ";" + "Extended Properties=Excel 8.0";
            }
            else
            {
                throw new Exception("文件格式有誤");
            }
            OleDbConnection oledbConn = new OleDbConnection(sConn);
            oledbConn.Open();
            OleDbDataAdapter command = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", oledbConn);
            DataSet ds = new DataSet();
            command.Fill(ds);
            oledbConn.Close();
            return ds.Tables[0];
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try{ mySQLConnection.Open();}
            catch (SqlException sex)
            {MessageBox.Show(sex.Message);}

            var tablename = new StringBuilder();
            tablename.Append("env");
            tablename.Append(filename);
            SqlCommand mySQLCommand0 = new SqlCommand();
            SqlTransaction tran;

            tran = mySQLConnection.BeginTransaction();


            try
            {
                for (int i = 1; i < dataGridView1.RowCount; i++)

                {


                    if (dataGridView1.Rows[i].Cells[1].Value != null)

                    {


                        mySQLCommand0 = new SqlCommand("INSERT INTO [" + tablename + "](編號, 裁處書字號, 裁罰對象, 違反法規, 裁處時數,違反日期,狀態,排課次數) VALUES (@id, @number, @who, @rule, @time ,@date,@state,@class)", mySQLConnection);

                        mySQLCommand0.Parameters.AddWithValue("@id", dataGridView1.Rows[i].Cells[0].Value.ToString());

                        mySQLCommand0.Parameters.AddWithValue("@number", dataGridView1.Rows[i].Cells[1].Value.ToString());

                        mySQLCommand0.Parameters.AddWithValue("@who", dataGridView1.Rows[i].Cells[2].Value.ToString());

                        mySQLCommand0.Parameters.AddWithValue("@rule", dataGridView1.Rows[i].Cells[3].Value.ToString());

                        mySQLCommand0.Parameters.AddWithValue("@time", dataGridView1.Rows[i].Cells[4].Value.ToString());

                        mySQLCommand0.Parameters.AddWithValue("@date", dataGridView1.Rows[i].Cells[5].Value.ToString());

                        mySQLCommand0.Parameters.AddWithValue("@state", dataGridView1.Rows[i].Cells[6].Value.ToString());

                        mySQLCommand0.Parameters.AddWithValue("@class", dataGridView1.Rows[i].Cells[7].Value.ToString());


                        mySQLCommand0.Transaction = tran;

                        mySQLCommand0.ExecuteNonQuery();



                    }
                }
            }
            catch (SqlException sex)
            {

            }

            
            MessageBox.Show("匯入成功");

            tran.Commit();
            mySQLConnection.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {

            this.Hide();

        }

        private void import_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }
    }
}
