using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace ExcelUpload.DAL
{
    public class DBConnection
    {
        public static SqlConnection getConnection()
        {
            String constr = ConfigurationManager.ConnectionStrings["dbconnection"].ToString();
            SqlConnection con = new SqlConnection(constr);
            return con;
        }
       
    }
}