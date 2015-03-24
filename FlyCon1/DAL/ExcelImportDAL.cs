using ExcelUpload.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;

namespace ExcelUpload.DAL
{
    public class ExcelImportDAL
    {

        #region Public Properties
        SqlConnection con = new SqlConnection();
        #endregion Public Properties

        #region Methods

        #region Method to get table definition
        public DataSet getTableDefinition(string TableName)
        {    
         
            SqlConnection con = new SqlConnection();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter da = new SqlDataAdapter();
            DataSet myRec = new DataSet();
            
            con = DAL.DBConnection.getConnection();           
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "procTableSelect";
            cmd.Parameters.Add("@tablename", SqlDbType.NVarChar).Value = TableName;// constantList.TableName;
            cmd.Connection = con;
            da.SelectCommand = cmd;
            
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
                da.Fill(myRec);
                return myRec;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                con.Close();
                con.Dispose();
            }
            
        }
        #endregion Method to get table definition

        #region Method to get Procedure Name
        public string getProcedureName(string TableName)
        {
            string procName;
            DAL.Constants constantList = new Constants();
            SqlConnection con = new SqlConnection();
            SqlCommand cmd = new SqlCommand();
            con = DAL.DBConnection.getConnection();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "procProcedureSelect";
            cmd.Parameters.Add("@tablename", SqlDbType.NVarChar).Value = TableName;// constantList.TableName;
            cmd.Parameters.Add("@propertyName", SqlDbType.NVarChar).Value = constantList.ImportProperty;
            cmd.Connection = con;
          
            try
            {
                con.Open();
                procName=cmd.ExecuteScalar().ToString();
                return procName;
         
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                con.Close();
                con.Dispose();
            }
        }
        #endregion Method to get Procedure Name

        #region Inserting Data From Dataset to Database

        /// <summary>
        /// Insert Excel File Datas in Dataset to Database
        /// </summary>
        /// <param name="dsFile"></param>
        /// <returns>success or failure</returns>
        public int InsertExcelFile(DataSet dsTable, DataRow dr, string excelFileName)
        {

            SqlCommand cmd = new SqlCommand();
            DAL.ExcelImportDAL stdDal = new DAL.ExcelImportDAL();
            DAL.Constants constantList = new DAL.Constants();
           
            dsTable = stdDal.getTableDefinition(constantList.TableName);

            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = stdDal.getProcedureName(constantList.TableName);
            cmd.Connection = con;
           
            for (int j = 0; j < dsTable.Tables[0].Rows.Count; j++)
            {
                string paramName = dsTable.Tables[0].Rows[j]["Field_Name"].ToString();
                string type = dsTable.Tables[0].Rows[j]["Field_DataType"].ToString();
                object paramValue = dr[paramName];

                if (type == "D")
                {
                    cmd.Parameters.AddWithValue(paramName, Convert.ToDateTime(paramValue));
                }
                else
                    cmd.Parameters.AddWithValue(paramName, paramValue);
            }
            cmd.Parameters.AddWithValue("@Updated_By", constantList.User);
            cmd.Parameters.AddWithValue("@Updated_Date", DateTime.Now);
            SqlParameter outPutParameter = new SqlParameter();
            outPutParameter.ParameterName = "@isUpdate";
            outPutParameter.SqlDbType = System.Data.SqlDbType.Int;
            outPutParameter.Direction = System.Data.ParameterDirection.Output;
            cmd.Parameters.Add(outPutParameter);
            cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            string IsUpdate = outPutParameter.Value.ToString();
            if (IsUpdate == "0")
            {
                return 1;             
            }   
            return 0;
        }
       
        #endregion Inserting Excel data in Dataset to Database

        #region Connect Database
        /// <summary>
        /// Opens the database connection
        /// </summary>
        public void ConnectDB()
        {
            con = DAL.DBConnection.getConnection();
            con.Open();
        }
        #endregion Connect Database

        #region Disconnect Database
        /// <summary>
        /// Closes the database connection
        /// </summary>
        public void DisconnectDB()
        {
            con.Close();
        }
        #endregion Disconnect Database

        #endregion Methods
    
    }
}