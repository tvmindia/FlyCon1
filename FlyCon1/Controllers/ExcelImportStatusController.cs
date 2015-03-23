using ExcelUpload.DAL;
using ExcelUpload.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelUpload.Controllers
{
    public class ExcelImportStatusController : Controller
    {
        //
        // GET: /ImportStatus/

        public ActionResult Index(string id)
        {
            List<ExcelImportStatusModel> lmd = new List<ExcelImportStatusModel>();
            DataSet ds = new DataSet();
            ExcelImportDetailsDAL  detailsObj = new ExcelImportDetailsDAL();

            ds = detailsObj.getExcelImportDetails(id);
       
            //return View(ds.Tables[0]);
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                    lmd.Add(new ExcelImportStatusModel
                    {
                        StatusId = dr["Status_Id"].ToString(),
                        ProjNo = dr["ProjNo"].ToString(),
                        FileName = dr["File_Name"].ToString(),
                        TableName = dr["Table_Name"].ToString(),
                        TotalCount = Convert.ToInt32(dr["Total_Count"]),
                        InsertCount = Convert.ToInt32(dr["Insert_Count"]),
                        UpdateCount = Convert.ToInt32(dr["Update_Count"]),
                        ErrorCount = Convert.ToInt32(dr["Error_Count"]),
                        UserName = dr["User_Name"].ToString(),
                        LastUpdatedTime = Convert.ToDateTime(dr["Last_Updated_Time"]),
                        StartTime = Convert.ToDateTime(dr["Start_Time"]),
                        Status = dr["StatusDescription"].ToString(),
                        Remarks = dr["Remarks"].ToString(),
                        TimeRemaining = TimeSpan.FromMilliseconds(Convert.ToDouble(dr["Time_Remaining"])).ToString(@"hh\:mm\:ss"),
                        TimeElapsed = TimeSpan.FromMilliseconds(Convert.ToDouble(dr["Time_Elapsed"])).ToString(@"hh\:mm\:ss")
                        //TimeRemaining = Convert.ToDecimal(dr["Time_Remaining"])
                    });
                
       
            }
            return View(lmd);
        }


        //
        // GET: /ImportStatus/ErrorDetails

        public ActionResult ErrorDetails(string id)
        {
            List<ExcelImportErrorDetails> Errorlmd = new List<ExcelImportErrorDetails>();
            DataSet ds = new DataSet();
            ExcelImportDetailsDAL detailsObj = new ExcelImportDetailsDAL();

            ds = detailsObj.getExcelImportErrorDetails(id);
            //return View(ds.Tables[0]);
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                Errorlmd.Add(new ExcelImportErrorDetails
                {
                   keyField = dr["Key_Field"].ToString(),
                   ErrorDescription = dr["Error_Description"].ToString()

                });
            }
            return View(Errorlmd);
        }

        public ActionResult AllExcelImports(string id)
        {           
           
            List<ExcelImportStatusModel> lmd = listDetails(id);       
            return View(lmd);
        }

        public ActionResult DeleteExcelImportDetails(string id)
        {
           
            ExcelImportDetailsDAL detailsObj = new ExcelImportDetailsDAL();
            int res = detailsObj.DeleteExcelImportDetails(id);
           // List<ExcelImportStatusModel> lmd = listDetails(id);
            return RedirectToAction("AllExcelImports");
        }

        public  List<ExcelImportStatusModel> listDetails(string id)
        {
            DataSet ds = new DataSet();
            ExcelImportDetailsDAL detailsObj = new ExcelImportDetailsDAL();
            ds = detailsObj.getExcelImportDetails(id);
            List<ExcelImportStatusModel> lmd = new List<ExcelImportStatusModel>();
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                lmd.Add(new ExcelImportStatusModel
                {
                    StatusId = dr["Status_Id"].ToString(),
                    FileName = dr["File_Name"].ToString(),
                    StartTime = Convert.ToDateTime(dr["Start_Time"]),
                    Status = dr["StatusDescription"].ToString()
                });
            }
            return lmd;
        }

    }
}
 