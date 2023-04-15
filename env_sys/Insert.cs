using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Net;
using System.IO;

namespace env_sys
{
    public partial class Insert : Form
    {
        public Insert()
        {
            InitializeComponent();
        }



        SQL_Con sob1 = new SQL_Con();
        DataSet ds = new DataSet();
        SqlDataAdapter myda1 = null;
        StreamReader sr = null;
        SqlCommand com = null;

        string receive_year, get_year, path;

        float ID=0,TIME=0,CLASS=0;

        int tb_va,lb_va = 0,count=0;
       



        private void Insert_Load(object sender, EventArgs e)
        {


            path = @"C:\env_system\env_sys\env_sys\column\name.ini";

           
            //dymaic textbox & label
            /* int index = 0;
            TextBox TB = new TextBox();
            Label LB = new Label();
            for(int x=8;x<=data.RowCount;x++)
             {
                 LB.Location = new Point(12, 572 + lb_va);
                 LB.Size = new Size(156, 28);
                 LB.TextAlign = ContentAlignment.MiddleCenter;
                 LB.BackColor = System.Drawing.Color.BlanchedAlmond;
                 LB.Text = data.Rows[index].Cells[x].ToString();
                 LB.Name = "txt" + count;
                 TB.Location = new Point(184, 563 + tb_va);
                 TB.Size = new Size(258, 37);
                 tb_va += 60;
                 lb_va += 68;
                 this.Controls.Add(LB);
                 this.Controls.Add(TB);
                 count++;
             }*/
            receive_year = Form1.year_option;

            sob1.txt_bind(receive_year);

            get_year = sob1.con_year;


            sob1.connect();

            //data.DataSource = null;

            sob1.sql = "select * from [" + get_year + "]";
            
            com = new SqlCommand(sob1.sql, sob1.connect());


            myda1 = new SqlDataAdapter(sob1.sql, sob1.connect());

            myda1.Fill(ds, get_year.ToString());


            data.DataSource = ds.Tables[get_year.ToString()];

            


           
            


        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (Control c in Controls)
            {
                if (c is TextBox )
                    c.Text = "";
            }

            data.DataSource = null;
            label10.Text = "";
            update_data.Text = "";
       }

        private void button3_Click_1(object sender, EventArgs e)
        {



            ds.Clear();
            ID = float.Parse(id.Text);

            TIME = float.Parse(time.Text);

            CLASS = float.Parse(clas.Text);
            
            ///////////////////////////////////////////update

            try
            {





               
                sob1.connect();
                sob1.sql = "select * from "+ get_year ;

                myda1 = new SqlDataAdapter(sob1.sql, sob1.connect());
                sob1.data_update(get_year, ID, number.Text, who.Text, rule.Text, TIME, date.Text, state.Text, CLASS, phone.Text, address.Text);

                int i = int.Parse(ID.ToString());

                
                myda1.Fill(ds, get_year.ToString());


                 data.DataSource = ds.Tables[get_year.ToString()];
                 /* MessageBox.Show(get_year);
                 MessageBox.Show(ID.ToString());
                 MessageBox.Show(number.Text);
                 MessageBox.Show(who.Text);
                 MessageBox.Show(rule.Text);
                 MessageBox.Show(TIME.ToString());
                 MessageBox.Show(date.Text);
                 MessageBox.Show(state.Text);
                 MessageBox.Show(CLASS.ToString());
                 MessageBox.Show(phone.Text);
                 MessageBox.Show(address.Text);
                 */
                         
                 MessageBox.Show("更新資料完畢");

                 label10.Text = get_year.Replace("env", "") + "年度搜尋結果, 共" + (data.RowCount - 1).ToString() + "筆資料";

                for (int j = 0; j <= data.ColumnCount - 1; j++)
                {

                    data.Rows[i].Cells[j].Style.BackColor = Color.Red;

                }
                update_data.Text = "更新的資料:" + "\r\n" + "編號: "+data.Rows[i].Cells[0].Value+ "\r\n" + "裁處書編號: " + data.Rows[i].Cells[1].Value + "\r\n" + "裁罰對象: " + data.Rows[i].Cells[2].Value + "\r\n" + "違反法規: " + data.Rows[i].Cells[3].Value + "\r\n" + "違反時數: " + data.Rows[i].Cells[4].Value + "\r\n" + "違反日期: " + data.Rows[i].Cells[5].Value + "\r\n" + "狀態: " + data.Rows[i].Cells[6].Value + "\r\n" + "排課次數: " + data.Rows[i].Cells[7].Value + "\r\n" + "電話: " + data.Rows[i].Cells[8].Value + "\r\n" + "地址: " + data.Rows[i].Cells[9].Value;

               
            }
            catch (SqlException sex)
            {
                MessageBox.Show(sex.Message);

            }

            sob1.close();
        }
       /* private void send_data(float ID, string number, string who, string rule,float TIME, string date, string state, float CLASS, string phone, string  address)
        {

            return ID;
        }
        */
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            id.Text = data.CurrentRow.Cells[0].Value.ToString();
            number.Text = data.CurrentRow.Cells[1].Value.ToString();
            who.Text = data.CurrentRow.Cells[2].Value.ToString();
            rule.Text = data.CurrentRow.Cells[3].Value.ToString();
            time.Text = data.CurrentRow.Cells[4].Value.ToString();
            date.Text = data.CurrentRow.Cells[5].Value.ToString();
            state.Text = data.CurrentRow.Cells[6].Value.ToString();
            clas.Text = data.CurrentRow.Cells[7].Value.ToString();
            phone.Text = data.CurrentRow.Cells[8].Value.ToString();
            address.Text = data.CurrentRow.Cells[9].Value.ToString();
            //send_data(float.Parse(id.Text), number.Text, who.Text, rule.Text, float.Parse(time.Text), date.Text, state.Text,float.Parse(clas.Text), phone.Text,address.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //show_data
            


            
        }

        private void label1_Click(object sender, EventArgs e)
        {



     
        }

        private void id_TextChanged(object sender, EventArgs e)
        {

            /*SqlCommand com = null;
            sob1.connect();
            


            try
                {

                    int i = int.Parse(id.Text);
                    float ID = float.Parse(id.Text);
                  
                    sob1.sql = "select * from [" + get_year + "] where 編號 ='" + i + "' ";

                    com = new SqlCommand(sob1.sql, sob1.connect());


                    myda1 = new SqlDataAdapter(sob1.sql, sob1.connect());

                    myda1.Fill(ds, get_year.ToString());
                    data.DataSource = ds.Tables[get_year.ToString()];
                

            }
                catch (Exception ex)
                {
                //MessageBox.Show(ex.Message);
                  
                }


               


                 sob1.close();
                
                

               */
            
           
            


            
         




            
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void update_data_Click(object sender, EventArgs e)
        {

        }

        private void number_TextChanged(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click_1(object sender, EventArgs e)
        {
           

        }

        private void button1_Click(object sender, EventArgs e)
        {

            ///////////////////////////////////////////add

            sob1.connect();
            
            
            ID = float.Parse(id.Text);

            TIME = float.Parse(time.Text);

            CLASS = float.Parse(clas.Text);
            if (ID==0 || TIME == 0 || CLASS==0)
            {
                ID = 0;
                TIME = 0;
                CLASS = 0;
            }



            try
            {
                sob1.data_insert(get_year, ID, number.Text, who.Text, rule.Text, TIME, date.Text, state.Text, CLASS,phone.Text,address.Text);
                MessageBox.Show("資料新增完畢");      
            }catch(SqlException sex)
            {
                //MessageBox.Show(sex.Message);
                MessageBox.Show("重複資料，無法新增，若數據需修改麻煩更新資料");
            }
            sob1.close();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            this.Hide();
        }

        private void Insert_FormClosing(object sender, FormClosingEventArgs e)
        {
            
           




        }
    }
}
