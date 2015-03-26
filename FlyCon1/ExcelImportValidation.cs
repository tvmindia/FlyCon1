using ExcelUpload.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace ExcelUpload
{
    public class ExcelImportValidation
    {

        #region Public Properties

        public string errorMessage
        {
            get;
            set;
        }

        public DataTable ErrorDatatable
        {
            get;
            set;
        }

        public Guid status_Id
        {
            get;
            set;
        }

        #endregion Public Properties

        #region Methods

        #region Data Validation
        /// <summary>
        /// Validation of Excel Data
        /// </summary>
        /// <param name="datarow from the result dataset of the excel file"></param>
        /// <returns>success/failure and error datatable</returns>
        public int excelDatasetValidation(DataRow dr)
        {
            DataTable dtError = CreateErrorTable();
            DataSet dsError = new DataSet();
            DataSet dsTable = new DataSet();
            DAL.Constants constantList = new DAL.Constants();
            DAL.ExcelImportDAL stdDal = new DAL.ExcelImportDAL();
            DAL.ExcelImportDetailsDAL detailsDal = new DAL.ExcelImportDetailsDAL(status_Id);
            List<ExcelValidationModel> lmd = new List<ExcelValidationModel>();

            dsTable = stdDal.getTableDefinition(constantList.TableName);
            DataRow[] result = dsTable.Tables[0].Select("ExcelMustFields='Y'");
            DataRow[] keyFieldRow = dsTable.Tables[0].Select("Key_Field='Y'");
                      
                StringBuilder keyFieldLists = new StringBuilder();
                StringBuilder errorDescLists = new StringBuilder();
                bool flag = false;
                string keyField = GetInvalidKeyField(keyFieldRow,dr);
                string comma = "";
                foreach (var item in result)
                {
                    string FieldName = item["Field_Name"].ToString();
                    string FieldDataType = item["Field_DataType"].ToString();

                    if (dr[FieldName].ToString().Trim() == "" || dr[FieldName] == null)
                    {
                        
                        flag = true;
                        errorDescLists.Append(comma);
                        errorDescLists.Append(FieldName);
                        errorDescLists.Append(" Field Is Empty");
                        comma = ",";
                       
                    }

                    else if (FieldDataType == "D" && !ValidateDate(dr[FieldName].ToString()))
                    {
                        flag = true;
                        errorDescLists.Append(comma);
                        errorDescLists.Append(FieldName);
                        errorDescLists.Append("is Invalid");
                        comma = ",";
             
                    }

                    else if (FieldDataType == "A" && !isAlphaNumeric(dr[FieldName].ToString()))
                    {
                        flag = true;
                        errorDescLists.Append(comma);
                        errorDescLists.Append(FieldName);
                        errorDescLists.Append(" is invalid");
                        comma = ",";
                      
                    }
                    
                    else if (FieldDataType == "N" && !isNumber(dr[FieldName].ToString()))
                    {
                        flag = true;
                        errorDescLists.Append(comma);
                        errorDescLists.Append(FieldName);
                        errorDescLists.Append(" is invalid");
                        comma = ",";
                          
                    }
                    
                    else if (FieldDataType == "S" && !isAlpha(dr[FieldName].ToString()))
                    {
                        flag = true;
                        errorDescLists.Append(comma);
                        errorDescLists.Append(FieldName);
                        errorDescLists.Append(" is invalid");
                        comma = ",";
                    }

                }
              
                if (flag == true) 
                {
                    detailsDal.InsertExcelImportErrorDetails(keyField, errorDescLists.ToString());
                    return -1;
                }
                else return 1;
        }

        /// <summary>
        /// Create datatable for Error Descriptions
        /// </summary>
        /// <returns></returns>
        public DataTable CreateErrorTable()
        {
            DataTable dtTemp = new DataTable();
            dtTemp.Columns.Add(new DataColumn("KeyField", typeof(string)));
            dtTemp.Columns.Add(new DataColumn("ErrorDesc", typeof(string)));
            return dtTemp;
        }


        #region Date validation
        /// <summary>
        /// Date validation
       /// </summary>
       /// <param name="date"></param>
       /// <returns></returns>
        private bool ValidateDate(string date)
        {
            try
            {
                DateTime tempDate;
                tempDate = Convert.ToDateTime(date);
                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }

        #endregion Date validation

        #region Alphanumeric Validation
        /// <summary>
        /// Alphanumeric Validation
        /// </summary>
        /// <param name="strToCheck"></param>
        /// <returns></returns>
        private static bool isAlphaNumeric(string strToCheck)
        {
            Regex rg = new Regex(@"^[a-zA-Z0-9\s,]*$");
            if (rg.IsMatch(strToCheck))
                return true;
            else
                return false;
        }

        #endregion Alphanumeric Validation

        #region Numeric Validation
        /// <summary>
        /// Numeric Validation
        /// </summary>
        /// <param name="strToCheck"></param>
        /// <returns></returns>
        private bool isNumber(string strToCheck)
        {
            Regex rg = new Regex(@"^[0-9\s,]+$");
            if (rg.IsMatch(strToCheck))
                return true;
            else
                return false;
        }

        #endregion Numeric Validation

        #region Alphabetic validation
        /// <summary>
        /// Alphabetic validation
        /// </summary>
        /// <param name="strToCheck"></param>
        /// <returns></returns>
        private bool isAlpha(string strToCheck)
        {
            Regex rg = new Regex(@"^[a-zA-Z\s,]+$");
            if (rg.IsMatch(strToCheck))
                return true;
            else
                return false;
        }
        
        #endregion Alphabetic validation

        #endregion Data Validation

        #region  ExcelDataStructureValidation
        /// <summary>
        /// Validation to check whether columns exist in excel file
        /// </summary>
        /// <param name="cName"></param>
        /// <param name="dsFile"></param>
        /// <returns></returns>
        public bool ExcelDataStructureValidation(string cName, DataSet dsFile)
        {
            for (int i = 0; i < dsFile.Tables[0].Columns.Count; i++)
            {
                if (cName == dsFile.Tables[0].Columns[i].ColumnName.ToString())
                    return true;
            }

            return false; 
        }
       

       
        public bool ValidateExcelDataStructure(DataSet dsFile)
        {
            DataSet dsTable = new DataSet();
            DAL.Constants constantList = new DAL.Constants();
            DAL.ExcelImportDAL stdDal = new DAL.ExcelImportDAL();
            dsTable = stdDal.getTableDefinition(constantList.TableName);
            for (int i = dsTable.Tables[0].Rows.Count - 1; i >= 0; i--)
            {
                bool res;
                string FieldName = dsTable.Tables[0].Rows[i]["Field_Name"].ToString();
                res = ExcelDataStructureValidation(FieldName,dsFile);
                if (!res)
                {
                    return false;
                }
            }

            return true;
        }

        #endregion  ExcelDataStructureValidation

        #region ValidateFileExtension
        public int ValidateFileExtension(string fileExtension) 
        {
            if (fileExtension != ".xls" && fileExtension != ".xlsx")
            {
                errorMessage = "Invalid File!";
                return 0;
            }
            return 1;

        }

        #endregion ValidateFileExtension

        #region GetInvalidKeyField
        private string GetInvalidKeyField(DataRow[] keyFieldRow,DataRow dr)
        {
            string comma = "";
            string keyField = "";
        
                for (int j = 0; j < keyFieldRow.Count(); j++)
                {
                    keyField += comma+ dr[keyFieldRow[j]["Field_Name"].ToString()].ToString() ;
                    comma = ",";
                }

            return keyField;
          }
        
        #endregion GetInvalidKeyField

        #endregion Methods
        //test
    }

}