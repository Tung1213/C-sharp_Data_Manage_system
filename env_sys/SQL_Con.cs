using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace env_sys
{
    class SQL_Con
    {

         public string sql,condition,path,con_year=null,format=",", sqlnew;
         public string remote_con,severaddress = "26.138.171.164", dbName = "Env_Data", userName = "user", password = "1234";
         private SqlConnection sc = null;
         private SqlCommand command = null;
        private StreamReader sr = null;
        
       

        public SqlConnection connect()
        {
            remote_con = "server=" + severaddress + ";database=" + dbName + ";Uid=" + userName + ";Pwd=" + password;
            sc = new SqlConnection(remote_con);
            sc.Open();
            return sc;
        
        }
        public void close()
        {
            sc.Close();
        }
        public void txt_bind(string year)
        {
            var content = new StringBuilder();
            content.Append("env");
            content.Append(year);
            con_year = content.ToString();
        }
        public void data_update(string get_year, float ID, string number, string who, string rule, float TIME, string date, string state, float CLASS,string Phone,string Address)
        {

            
           
            
            sql = "UPDATE [" + get_year + "] SET 編號=@Column1,裁處書字號=@Column2,裁罰對象=@Column3,違反法規=@Column4,裁處時數=@Column5,違反日期=@Column6,狀態=@Column7,排課次數=@Column8,電話=@Column9,地址=@Column10 where 編號=@Column1";
          
            command = new SqlCommand(sql.ToString(), sc);
            command.Parameters.Add("@Column1", SqlDbType.Float).Value = ID;
            command.Parameters.Add("@Column2", SqlDbType.NVarChar).Value = number;
            command.Parameters.Add("@Column3", SqlDbType.NVarChar).Value = who; ;
            command.Parameters.Add("@Column4", SqlDbType.NVarChar).Value = rule;
            command.Parameters.Add("@Column5", SqlDbType.Float).Value = TIME;
            command.Parameters.Add("@Column6", SqlDbType.NVarChar).Value = date;
            command.Parameters.Add("@Column7", SqlDbType.NVarChar).Value = state;
            command.Parameters.Add("@Column8", SqlDbType.Float).Value = CLASS;
            command.Parameters.Add("@Column9", SqlDbType.NVarChar).Value = Phone;
            command.Parameters.Add("@Column10", SqlDbType.NVarChar).Value = Address;

            command.ExecuteNonQuery();
            
            
       
        }
        public bool data_insert(string get_year,float ID, string number,string who,string rule,float TIME,string date,string state,float CLASS, string Phone, string Address)
        {
          
            sql = "insert into [" + get_year + "] (編號,裁處書字號,裁罰對象,違反法規,裁處時數,違反日期,狀態,排課次數) VALUES(@Column1,@Column2,@Column3,@Column4,@Column5,@Column6,@Column7,@Column8,@Column8,@Column10)";
            command = new SqlCommand(sql,sc);             
            command.Parameters.Add("@Column1", SqlDbType.Float).Value = ID;
            command.Parameters.Add("@Column2", SqlDbType.NVarChar).Value = number;
            command.Parameters.Add("@Column3", SqlDbType.NVarChar).Value = who;
            command.Parameters.Add("@Column4", SqlDbType.NVarChar).Value = rule;
            command.Parameters.Add("@Column5", SqlDbType.Float).Value = TIME;
            command.Parameters.Add("@Column6", SqlDbType.NVarChar).Value = date;
            command.Parameters.Add("@Column7", SqlDbType.NVarChar).Value = state;
            command.Parameters.Add("@Column8", SqlDbType.Float).Value = CLASS;
            command.Parameters.Add("@column9", SqlDbType.NVarChar).Value = Phone;
            command.Parameters.Add("@column10", SqlDbType.NVarChar).Value = Address;
            command.ExecuteNonQuery();
          
            return true;
        }
        

    }
}
