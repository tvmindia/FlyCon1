using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Threading;
using ExcelUpload.DAL;
using ExcelUpload.Models;
namespace ExcelUpload.Controllers
{
    public class ExcelImportController : Controller
    {
        //
        // GET: /ExcelImport/

        UpdatedExcelImport importObj = new UpdatedExcelImport();
        Constants constantsObj = new Constants();
        public ActionResult Index()
        {
            return View();

        }
       
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            UpdatedExcelImport importObj = new UpdatedExcelImport();

            //int result;
            importObj.tableName = constantsObj.TableName;
            importObj.fileName = "file";
            importObj.request = Request;
            importObj.temporaryFolder = Server.MapPath("~/Content/").ToString();
            
            Thread excelImportThread = new Thread(new ThreadStart(importObj.ImportExcelFile));
            excelImportThread.Start();

            if (importObj.importStatus == -1)
                {
                    TempData["Message"] = importObj.successMessage;

                }

            return View();
        }
    }
}
