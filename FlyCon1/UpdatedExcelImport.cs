using ExcelUpload.DAL;
using ExcelUpload.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace ExcelUpload
{
    public class UpdatedExcelImport
    {
        #region Public Properties
        ExcelImportValidation validationObj = new ExcelImportValidation();
        public string errorMessage
        {
            get;
            set;
        }
        public string successMessage
        {
            get;
            set;
        }
        public string failureMessage
        {
            get;
            set;
        }
        public List<ExcelValidationModel> invalidItems
        {
            get;
            set;
        }
        public int insertcount
        {
            get;
            set;
        }
        public int totalCount
        {
            get;
            set;
        }
        public int updateCount
        {
            get;
            set;
        }

        public int errorCount
        {
            get;
            set;
        }

        public string remarks
        {
            get;
            set;
        }

        public string tableName
        {
            get;
            set;
        }

        public string fileName
        {
            get;
            set;
        }

        public int importStatus
        {
            get;
            set;
        }

        public HttpRequestBase request
        {
            get;
            set;
        }
        public string temporaryFolder
        {
            get;
            set;
        }
        public string ExcelFileName
        {
            get;
            set;
        }

 

        #endregion Public Properties

        #region Methods

        #region ImportExcelFile and Validation
        /// <summary>
        /// Import and validate excel file.
        /// </summary>

        public void ImportExcelFile()
        {
            #region Properties
            
            var Request = request;
            string tempFolder = temporaryFolder;
            //string tempFolder = Path.Combine(HttpRuntime.AppDomainAppPath, "~/Content/");
            DataSet dsFile = new DataSet();
            //DataTable dtError;
            List<ExcelValidationModel> lmd = new List<ExcelValidationModel>();

            #endregion Properties

            #region Reading Excel File To Dataset
            if (Request.Files[fileName].ContentLength > 0)
            {
                int fileExtensionCheck;
                string fileExtension = System.IO.Path.GetExtension(Request.Files[fileName].FileName);
                ExcelFileName = Request.Files[fileName].FileName;
                
                fileExtensionCheck=validationObj.ValidateFileExtension(fileExtension);
                
                if(fileExtensionCheck == 0)
                {
                    importStatus = -1;
                    return;
                }

                else
                {
                    string fileLocation = tempFolder + Request.Files[fileName].FileName;
                    if (System.IO.File.Exists(fileLocation))
                    {
                        try
                        {
                            System.IO.File.Delete(fileLocation);
                        }

                        catch
                        {
                            errorMessage = "Please try again!";
                            importStatus = -1;
                            return;
                        }
                    }

                    Request.Files[fileName].SaveAs(fileLocation);
                    string excelConnectionString = string.Empty;

                    if (fileExtension == ".xls")
                    {
                        excelConnectionString = System.Configuration.ConfigurationManager.AppSettings["XLS_ConnectionString"];
                        excelConnectionString = excelConnectionString.Replace("$fileLocation$", fileLocation);
                    }
                    //connection String for xlsx file format.
                    else if (fileExtension == ".xlsx")
                    {
                        excelConnectionString = System.Configuration.ConfigurationManager.AppSettings["XLSX_ConnectionString"];
                        excelConnectionString = excelConnectionString.Replace("$fileLocation$", fileLocation);
                    }

                    //Create Connection to Excel work book and add oledb namespace
                    OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
                    excelConnection.Open();
                    DataTable dt = new DataTable();

                    dt = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dt == null)
                    {
                        importStatus = -1;
                        return;
                    }

                    String[] excelSheets = new String[dt.Rows.Count];
                    int t = 0;
                    //excel data saves in temp file here.
                    foreach (DataRow row in dt.Rows)
                    {
                        excelSheets[t] = row["TABLE_NAME"].ToString();
                        t++;
                    }
                    OleDbConnection excelConnection1 = new OleDbConnection(excelConnectionString);
                    string query = string.Format("Select * from [{0}]", excelSheets[0]);
                    using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, excelConnection1))
                    {
                        dataAdapter.Fill(dsFile);
                        excelConnection.Close();

                        totalCount = dsFile.Tables[0].Rows.Count;

                        if (dsFile.Tables[0].Rows.Count == 0)
                        {
                            failureMessage = "No data found!";
                            importStatus = -1;
                            return;
                        }
                    }

                }

            #endregion Reading Excel File To Dataset

            int result = InsertExcelFile(dsFile);
                                
            if (result == -1)
            {
               errorMessage = "Invalid Excel";
               importStatus = -1;
               return;                
              }      
            }
            invalidItems = lmd;
            importStatus = 1;
            return;
        }

        #region Inserting Data From Dataset to Database

        /// <summary>
        /// Insert Excel File Datas in Dataset to Database
        /// </summary>
        /// <param name="dsFile"></param>
        /// <returns>success or failure</returns>
        private int InsertExcelFile(DataSet dsFile)
        {
           
            DataTable dtError = validationObj.CreateErrorTable();
            DataSet dsTable = new DataSet();
            DAL.Constants constantList = new DAL.Constants();
            DAL.ExcelImportDAL importDal = new DAL.ExcelImportDAL();   
            ExcelImportDetailsDAL importDetailsObj = new ExcelImportDetailsDAL();
         
            dsTable = importDal.getTableDefinition(constantList.TableName);
            DataRow[] result = dsTable.Tables[0].Select("ExcelMustFields='Y'");
            DataRow[] keyFieldRow = dsTable.Tables[0].Select("Key_Field='Y'");

            validationObj.status_Id = importDetailsObj.status_Id;
            bool columnExistCheck = validationObj.ValidateExcelDataStructure(dsFile);

            if (columnExistCheck == false)
            {
                return -1;
            }

            importDetailsObj.InitializeExcelImportDetails(ExcelFileName,totalCount);

            importDal.ConnectDB();

            for (int i = dsFile.Tables[0].Rows.Count - 1; i >= 0; i--)
            {

                StringBuilder keyFieldLists = new StringBuilder();
                StringBuilder errorDescLists = new StringBuilder();
                int res;

                res = validationObj.excelDatasetValidation(dsFile.Tables[0].Rows[i]);
                if (res == -1)
                {
                    errorCount = errorCount + 1;
                }
                else if (res == 1)
                {
                    int insertResult;
                    insertResult= importDal.InsertExcelFile(dsTable, dsFile.Tables[0].Rows[i], ExcelFileName);
                    if (insertResult == 1)
                    {
                        insertcount = insertcount + 1;
                    }
                    else if (insertResult == 0)
                    {
                        updateCount = updateCount + 1;
                    }
                }
                
                importDetailsObj.UpdateExcelImportDetails(ExcelFileName, insertcount, updateCount, errorCount, remarks, excelImportstatus.Processing);
            }

            importDetailsObj.UpdateExcelImportDetails(ExcelFileName, insertcount, updateCount, errorCount, remarks, excelImportstatus.Finished);
            importDal.DisconnectDB();
            return 1;
        }

        #endregion Inserting Excel data in Dataset to Database

        #endregion ImportExcelFile and Validation

      
        #endregion Methods

    }
}