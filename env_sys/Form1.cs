using Microsoft.Win32;
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
using System.Resources;
using ExcelDataReader;
using System.Net;

namespace env_sys
{
    public partial class Form1 : Form
    {

        public static string year_option = null;
        public static string newcolumn = null;

        private Bitmap myimage;
        public Form1()
        {
            InitializeComponent();
        }

      
        string[] year = new string[1000];
        string path = null;
        int triger;

        string out_con = null, severaddress = "26.138.171.164", dbName = "Env_Data", userName = "user", password = "1234";
        //247
        /*string ConnectionString = "Data Source=192.168.43.247\\DESKTOP-J2VHBBU,1433;user id= sa ;password=1234 ;" +
                "Initial Catalog=Env_Data;" +
                "Integrated Security=SSPI;";
        */
        SqlConnection sn = null;

        SqlCommand cmd = null;




        //------------------------------       
        //---------------------------


        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

            string hostname = Dns.GetHostName();

            string LocalIP = Dns.GetHostByName(hostname).AddressList[0].ToString();

            severaddress = "26.138.171.164";
            //MessageBox.Show(severaddress);

            path = @"C:\env_system\env_sys\env_sys\combobox_option\Option.ini";

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 15);

            out_con = "server=" + severaddress + ";database=" + dbName + ";Uid=" + userName + ";Pwd=" + password;

            sn = new SqlConnection(out_con);
            
            try
            {
                if (File.Exists(path))
                {
                    StreamReader sr = new StreamReader(path, Encoding.Default);

                    while (sr.Peek() > 0)
                    {
                        
                        comboBox1.Items.Add(sr.ReadLine());
                        //year[count]=comboBox1.SelectedText.ToString();
                        //count++;
                    }
                    sr.Close();
                }
                //MessageBox.Show(comboBox1.Items.Count.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



            //////////////////////////////////////////////////////


            for (int x = 0; x < 101; x++) if (year[x] != null) { comboBox1.Items.Add(year[x]); } else break;
            comboBox1.SelectedItem = year[0];
            toolTip1.SetToolTip(button1, "請先選擇要查詢年度");
            toolTip2.SetToolTip(button4, "ex:主索引名稱");
            toolTip3.SetToolTip(button2, "請先輸入要匯入的年度");
            toolTip4.SetToolTip(button7, "請輸入編號(僅限數字)!!");
            toolTip5.SetToolTip(button10, "先確認要新增資料的年度");
            button4.Enabled = false;
            //pictureBox1.Image = Image.FromFile("C:\\env_system\\env_sys\\env_sys\\image\\env1.PNG");
            pictureBox2.Visible = true;
            //showiage("C:\\env_system\\env_sys\\env_sys\\image\\welcome.gif", 291, 94);
            //excel_grid();

        }
        //---------------------------------------------- slef-created method 
        private void showiage(String fileToDisplay, int xSize, int ySize)
        {
            if (myimage != null)
            {
                myimage.Dispose();
            }
            pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
            myimage = new Bitmap(fileToDisplay);
            pictureBox2.ClientSize = new Size(xSize, ySize);
            pictureBox2.Image = (Image)myimage;


        }
        //----------------------------------------------

        private void button1_Click(object sender, EventArgs e)
        {         
            connect_sql_server(comboBox1.SelectedIndex.ToString(), "", "");

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //匯入        
            try
            {
                if (textBox5.Text == "")
                {
                    MessageBox.Show("輸入年度");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

 


        /////////////////////////////////////////////////// Add_Dropdownlist_option
        private void add_dropdownlist(int row, string newyear)
        {


            if (comboBox1.Items.Contains(textBox5.Text))
            {
                MessageBox.Show("無法加入此年度，已存在這年度!");
            }
            else
            {
                comboBox1.Items.Add(newyear);
            }
            //儲存combobox的選項內容到配置檔案1.ini
            StreamWriter sw = new StreamWriter(path);
            for (int i = 0; i < comboBox1.Items.Count; i++)
            {
                sw.WriteLine(comboBox1.Items[i]);
            }
            sw.Close();




        }


        ///////////////////////////////////////////////////


        /*private void button3_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|文字檔 (Tab 字元分隔) (*.txt)|*.txt";
                try
                {
                    ofd.ShowDialog();


                }
                catch
                {
                    textBox1.Text = "error" + string.Empty;

                }
                finally
                {
                    textBox1.Text = "檔案開啟成功,檔名為" + ofd.FileName;
                }

            }
        }*/



        //------------------------------------------------------------------------

        // conect sql server and search keyword
        private void connect_sql_server(string index, string key_txt, string name_txt)
        {


                //Value.GetType().Name
                string sql;
                
                DataSet myds1 = new DataSet();

                sn.Open();
                var str = new StringBuilder();
                str.Append("env");
                str.Append(comboBox1.SelectedItem);
           
           
                if (key_txt == "" && name_txt == "" || button4.Enabled == false)
                {

                    sql = "select * from " + str;
                    SqlDataAdapter myda1 = new SqlDataAdapter(sql, sn);
                    myda1.Fill(myds1, str.ToString());
                    dataGridView1.DataSource = myds1.Tables[str.ToString()];
                    richTextBox1.Text = "目前資料:" + comboBox1.SelectedItem.ToString() + "年度,共" + (dataGridView1.RowCount - 1).ToString() + "筆資料";
              

            }

            else
                {
                    try
                    {
                    sql = string.Format("select * from [" + str + "] where [" + key_txt + "] LIKE \'%" + (name_txt + "%\'"));//[" + key_txt + "] = '" + name_txt + "'

                        SqlDataAdapter myda2 = new SqlDataAdapter(sql, sn);
                        //DataSet myds2 = new DataSet();
                        myda2.Fill(myds1, str.ToString());
                        dataGridView1.DataSource = myds1.Tables[str.ToString()];
                        richTextBox1.Text = "目前資料:" + comboBox1.SelectedItem.ToString() + "年度 " + "'" + key_txt + "'" + "索引查詢結果,共"+(dataGridView1.RowCount - 1).ToString() +"筆資料";
                    }
                    catch (System.Data.SqlClient.SqlException sqlException)
                    {
                        System.Windows.Forms.MessageBox.Show(sqlException.Message);
                    }
                }

            

            sn.Close();



        }

        //----------------------------------------------

        private void exportexcel(DataGridView dg1)
        {
   
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int i = 0;
            int j = 0;

            for (i = 0; i <= dataGridView1.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dataGridView1[j, i];
                    xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
                }
            }

            xlWorkBook.SaveAs("output.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("檔案建立，存在文件檔案裡");       
        }

        //----------------------------------------------
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        //----------------------------------------------
        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            richTextBox1.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string txt_index = comboBox2.SelectedItem.ToString();
            string txt_search_name = textBox3.Text.ToString();
            if (txt_index == "" && txt_search_name == "") MessageBox.Show("請輸入!!!");
            else if (txt_index == "" || txt_search_name == "")
            {
                MessageBox.Show("請都要輸入!!!");
            }
            else if (txt_index != "" && txt_search_name != "")
            {
                connect_sql_server(comboBox1.SelectedIndex.ToString(), txt_index, txt_search_name);

            }


        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            button4.Enabled = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            button4.Enabled = false;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            

        }

        private void button7_Click(object sender, EventArgs e)
        {

            if (textBox3.Text.ToString() == "" && textBox4.Text.ToString() == "")
            {
                MessageBox.Show("請輸入!!!");
            }
            else if (textBox3.Text.ToString() == "" || textBox4.Text.ToString() == "")
            {
                MessageBox.Show("請全部輸入!!!");
            }
            else
            {

                //do something ...
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            //textBox1.Text = "";
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {



            exportexcel(dataGridView1);

        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click_1(object sender, EventArgs e)
        {


            int id;
            try
            {
                id = int.Parse(textBox4.Text);
                delete_data(id, comboBox1.SelectedItem.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }




        }

        //1.不確定是否要用編號當作刪除key
        //2.判斷是否有被刪除
        private void delete_data(float id, string com)
        {
          
            var str = new StringBuilder();
            str.Append("env");
            str.Append(com);
            //sn = new SqlConnection(out_con);
            sn.Open();

            string sqlStatement = "UPDATE [" + str + "] SET 狀態='已完成' where 編號 = '" + id + "' ";

            try
            {
                cmd = new SqlCommand(sqlStatement, sn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("結案資料完畢!!");
            }
            catch (SqlException ex)
            {

                MessageBox.Show(ex.Message);


            }

            sn.Close();

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {


        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {
            //textBox1.Text = "";
        }

        private void button11_Click(object sender, EventArgs e)
        {


            int row_data = comboBox1.Items.Count;
            string new_year = textBox5.Text.ToString();
            add_dropdownlist(row_data, new_year);
        }


        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            //年度防止重複機制(有點問題)
            /*bool check_year = false;
           
            for (int i = 0; i < year.Length; i++) if (textBox5.Text.ToString() == year[i]) check_year = true;
            if (check_year == true) { MessageBox.Show("有重複的年度唷!!"); }
            */
        }

        ///////////////////////////////////////////////D
        private void button13_Click(object sender, EventArgs e)
        {


            //測試用
            try
            {
               
            }
            catch (Exception ex)
            {


            }      

        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show("謝謝使用本系統!!");
            Application.ExitThread();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            connect_sql_server(comboBox1.SelectedIndex.ToString(), "", "");
            
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            richTextBox1.Text = "";
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            string txt_index = comboBox2.SelectedItem.ToString();
            string txt_search_name = textBox3.Text.ToString();
            if (txt_index == "" && txt_search_name == "") MessageBox.Show("請輸入!!!");
            else if (txt_index == "" || txt_search_name == "")
            {
                MessageBox.Show("請都要輸入!!!");
            }
            else if (txt_index != "" && txt_search_name != "")
            {
                connect_sql_server(comboBox1.SelectedIndex.ToString(), txt_index, txt_search_name);

            }
        }

        private void radioButton1_CheckedChanged_1(object sender, EventArgs e)
        {
            button4.Enabled = true;
        }

        private void radioButton2_CheckedChanged_1(object sender, EventArgs e)
        {
            button4.Enabled = false;
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            exportexcel(dataGridView1);
        }

      /* private void button3_Click_1(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|文字檔 (Tab 字元分隔) (*.txt)|*.txt";
                try
                {
                    ofd.ShowDialog();
                    import_filename = ofd.FileName;

                }
                catch
                {
                    textBox1.Text = "error" + string.Empty;

                }
                finally
                {
                    textBox1.Text = "檔案開啟成功,檔名為" + ofd.FileName;
                }

            }
        }
        */
        ///////////////////////////////////////////
        private void prevent_add(string txt,int total,string new_year)
        {


            if (textBox5.Text.Trim() == "")
            {
                MessageBox.Show("輸入年度");
                
            }
            else if (comboBox1.Items.Contains(textBox5.Text))
            {
                MessageBox.Show("已存在這年度囉");
            }
            else
            {
                if (triger == 1)
                {
                    add_dropdownlist(total,new_year);
                }               
            }

        }
        


        ///////////////////////////////////////////



        private void button11_Click_1(object sender, EventArgs e)
        {
            int row_data = comboBox1.Items.Count;
            string new_year = textBox5.Text.ToString();
            triger = 1;
            prevent_add(textBox5.Text.ToString(),row_data,new_year);
           
            
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            import im = new import();
            Form1 f1 = new Form1();
            if (textBox5.Text == "")
            {
                MessageBox.Show("請選擇年度");

            }
            else
            {

            try
            {

                create_db_table(textBox5.Text);
                f1.Close();
                im.Show();
                
                   
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            }
            sn.Close();



        }

        //------------------------------------------------------
        private void create_db_table(string year)
        {


            ////////////////////////////////////////////////////create table
            var str = new StringBuilder();
            str.Append("env");
            str.Append(year);
            SqlDataReader reader = null;
           
            //SqlCommand cmd = null;
            sn = new SqlConnection();
            

            string sql;
          
            sn.ConnectionString = out_con;

            try
            {
                sn.Open();

            }
            catch (SqlException ex)
            {
                //MessageBox.Show(ex.Message);
            }

            sql = "create table [" + str + "]"+ "(編號 float primary key," + " 裁處書字號 nvarchar(255), 裁罰對象 nvarchar(255), 違反法規 nvarchar(255),裁處時數 float, 違反日期 nvarchar(255), 狀態 nvarchar(255),排課次數 float,電話 nvarchar(255),地址 nvarchar(255))";

            cmd = new SqlCommand(sql, sn);
            try
            {
                cmd.ExecuteNonQuery();
                error_show.Text = "匯入完畢";

        
            }catch(SqlException ex)
            {
                //error_show.Text = ex.Message.ToString();
                error_show.Text = "建立完畢，檔案儲存完畢";
            }




          


        }

     

        private void button7_Click_2(object sender, EventArgs e)
        {

            int id;
           
            
                try
                {
                    id = int.Parse(textBox4.Text);
                    delete_data(id, comboBox1.SelectedItem.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            
            
        }

       /* private void button13_Click_1(object sender, EventArgs e)
        {
            //測試用
            SqlConnection conn = new SqlConnection();
            try
            {
                
                conn.ConnectionString =
                "Data Source=192.168.43.247\\DESKTOP-J2VHBBU,1433;user id= sa ;password=1234 ;" +
                "Initial Catalog=Env_Data;" +
                "Integrated Security=SSPI;"
                
                ;
                conn.Open();
                MessageBox.Show("連線成功!!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            conn.Close();

        }*/

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            int current_remove_item = comboBox1.SelectedIndex;

            //remove_item(current_remove_item);

           

            



        }
        //----------------------------------------------

        /*private void remove_item(int index)
        {

            int current_item = comboBox1.Items.Count;
            StreamWriter sw = new StreamWriter(path);
            comboBox1.Items.RemoveAt(index);
            for(int x = 0; x < comboBox1.Items.Count; x++)
            {

            }

        


        }*/

        private void groupBox1_Enter_1(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            Insert in1 = new Insert();
            Form1 f = new Form1();
            try
            {
                year_option = comboBox1.SelectedItem.ToString();
                
                f.Show();
                this.Hide();
                in1.Visible = true;

            }
            catch(Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("請先選擇要修改資料的年度");
            }
           
           
        }

        private void groupBox2_Enter_1(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            textBox5.Text = comboBox1.SelectedItem.ToString();
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button13_Click_2(object sender, EventArgs e)
        {
          
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
           

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged_1(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //外機連線 ConnectionString

             try
            {

                sn = new SqlConnection(out_con);
                sn.Open();
                MessageBox.Show("ok");

            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            sn.Close();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged_1(object sender, EventArgs e)
        {

        }


        //----------------------------------------------





    }
}
