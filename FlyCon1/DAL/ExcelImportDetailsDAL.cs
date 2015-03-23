using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace ExcelUpload.DAL
{
    public class ExcelImportDetailsDAL
    {

        #region Initialize Excel Import Details

        public Guid status_Id
        {
            get;
            set;
        }

        public ExcelImportDetailsDAL()
        {

            status_Id = Guid.NewGuid();
        
        }

        public ExcelImportDetailsDAL(Guid StatusId)
        {

            status_Id = StatusId;

        }
        /// <summary>
        /// Initialize the values of Excel import details table 
        /// </summary>
        /// <param name="ExcelFileName"></param>
        public void InitializeExcelImportDetails(string ExcelFileName,int totalCount)
        {
            SqlConnection con = new SqlConnection();
            SqlCommand cmd = new SqlCommand();
            DAL.Constants constantList = new DAL.Constants();
            UpdatedExcelImport fileObj = new UpdatedExcelImport();
           
            con = DAL.DBConnection.getConnection();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "procExcelImportDetailsInsert";
            cmd.Connection = con;
            cmd.Parameters.AddWithValue("@Status_Id", status_Id);
            cmd.Parameters.AddWithValue("@ProjNo",constantList.ProjNo);
            cmd.Parameters.AddWithValue("@File_Name", ExcelFileName);
            cmd.Parameters.AddWithValue("@Table_Name",constantList.TableName);
            cmd.Parameters.AddWithValue("@Total_Count",totalCount);
            cmd.Parameters.AddWithValue("@Insert_Count",0);
            cmd.Parameters.AddWithValue("@Update_Count",0);
            cmd.Parameters.AddWithValue("@Error_Count",0);
            cmd.Parameters.AddWithValue("@User_Name",constantList.User);
            cmd.Parameters.AddWithValue("@InsertStatus",excelImportstatus.started);
            cmd.Parameters.AddWithValue("@Remarks","");
            //cmd.Parameters.AddWithValue("@Updated_Date",DateTime.Now);
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }

        #endregion Initialize Excel Import Details
       
        #region Update Excel Import Details
        /// <summary>
        /// Update the values of Excel import details table
        /// </summary>
        /// <param name="ExcelFileName"></param>
        /// <param name="InsertCount"></param>
        /// <param name="UpdateCount"></param>
        /// <param name="ErrorCount"></param>
        /// <param name="Remarks"></param>
        public void UpdateExcelImportDetails(string ExcelFileName,int InsertCount,int UpdateCount,int ErrorCount,string Remarks,excelImportstatus processStatus)
        {
            SqlConnection con = new SqlConnection();
            SqlCommand cmd = new SqlCommand();
            DAL.Constants constantList = new DAL.Constants();
                     
            con = DAL.DBConnection.getConnection();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "procExcelImportDetailsUpdate";
            cmd.Connection = con;
            cmd.Parameters.AddWithValue("@Status_Id", status_Id);
            cmd.Parameters.AddWithValue("@ProjNo", constantList.ProjNo);
            cmd.Parameters.AddWithValue("@File_Name", ExcelFileName);
            cmd.Parameters.AddWithValue("@Table_Name", constantList.TableName);
            cmd.Parameters.AddWithValue("@Insert_Count", InsertCount);
            cmd.Parameters.AddWithValue("@Update_Count", UpdateCount);
            cmd.Parameters.AddWithValue("@Error_Count", ErrorCount);
            cmd.Parameters.AddWithValue("@User_Name", constantList.User);
            cmd.Parameters.AddWithValue("@InsertStatus",processStatus);
            cmd.Parameters.AddWithValue("@Remarks",Remarks);
            //cmd.Parameters.AddWithValue("@Updated_Date", DateTime.Now);
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

        }
        #endregion Insert Excel Import Details

        #region getExcelImportDetails
        /// <summary>
        /// Gets values from Excel import details table
        /// </summary>
        /// <returns></returns>
        public DataSet getExcelImportDetails(string id)
        {
            SqlConnection con = new SqlConnection();
            SqlCommand cmd = new SqlCommand();
            DataSet myRec = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter();
            DAL.Constants constantList = new DAL.Constants();
            
            con = DAL.DBConnection.getConnection();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "procExcelImportDetailsSelect";
            cmd.Parameters.AddWithValue("@Status_Id", id);
            cmd.Parameters.AddWithValue("@User_Name",constantList.User);
            cmd.Connection = con;
            da.SelectCommand = cmd;
            con.Open();
            da.Fill(myRec);
            con.Close();
            return myRec;
        }     
        #endregion getExcelImportDetails

        #region Insert Excel Import Error Details
        /// <summary>
        /// Insert the values of Excel import error details table
        /// </summary>
        /// <param name="ExcelFileName"></param>
        /// <param name="InsertCount"></param>
     
        public void InsertExcelImportErrorDetails(string KeyField,string ErrorDescription)
        {
            SqlConnection con = new SqlConnection();
            SqlCommand cmd = new SqlCommand();
            DAL.Constants constantList = new DAL.Constants();

            con = DAL.DBConnection.getConnection();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "procExcelImportErrorDetailsInsert";
            cmd.Connection = con;
            cmd.Parameters.AddWithValue("@Import_Status_Id", status_Id);
            cmd.Parameters.AddWithValue("@Key_Field",KeyField);
            cmd.Parameters.AddWithValue("@Error_Description",ErrorDescription);
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

        }
        #endregion Insert Excel Import Error Details

        #region getExcelImportErrorDetails
        /// <summary>
        /// Gets values from Excel import error details table
        /// </summary>
        /// <returns></returns>
        public DataSet getExcelImportErrorDetails(string id)
        {
            SqlConnection con = new SqlConnection();
            SqlCommand cmd = new SqlCommand();
            DataSet myRec = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter();

            con = DAL.DBConnection.getConnection();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "procExcelImportErrorDetailsSelect";
            cmd.Parameters.AddWithValue("@Status_Id", id);
            cmd.Connection = con;
            da.SelectCommand = cmd;
            con.Open();
            da.Fill(myRec);
            con.Close();
            return myRec;
        }
        #endregion getExcelImportErrorDetails

        #region DeleteExcelImportDetails
        /// <summary>
        ///  Set delete flag of Excel import details table
        /// </summary>
        /// <returns></returns>
        public int DeleteExcelImportDetails(string id)
        {
            SqlConnection con = new SqlConnection();
            SqlCommand cmd = new SqlCommand();
            DAL.Constants constantList = new DAL.Constants();

            con = DAL.DBConnection.getConnection();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "procImportErrorDetailsDelete";
            cmd.Parameters.AddWithValue("@Status_Id", id);
            cmd.Parameters.AddWithValue("@User_Name", constantList.User);
            cmd.Connection = con;           
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
            return 1;
        }
        #endregion DeleteExcelImportDetails

    }
}