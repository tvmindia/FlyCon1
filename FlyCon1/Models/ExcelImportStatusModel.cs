using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelUpload.Models
{
    public class ExcelImportStatusModel
    {
        public string StatusId { get; set; }

        public string ProjNo { get; set; }

        public string FileName { get; set; }

        public string TableName { get; set; }
        
        public int TotalCount { get; set; }
        
        public int InsertCount { get; set; }

        public int UpdateCount { get; set; }

        public int ErrorCount { get; set; }

        public string UserName { get; set; }

        public DateTime LastUpdatedTime { get; set; }

        public DateTime StartTime { get; set; }

        public string Status { get; set; }

        public string Remarks { get; set; }

        public string TimeRemaining { get; set; }

        public string TimeElapsed { get; set; }
    }
}