using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelUpload.DAL
{
    public class Constants
    {
        public string TableName = "MASTER_Location";
        public string ImportProperty = "ImportProcedure";
        public string User = "TestUser";
        public string ProjNo = "CCMS001";
        public excelImportstatus ImportProcessState;
        
    }


    public enum excelImportstatus
    { 
       started = 1,
       Processing = 2, 
       Finished = 3
    }
}