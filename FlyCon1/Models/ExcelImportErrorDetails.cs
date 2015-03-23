using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelUpload.Models
{
    public class ExcelImportErrorDetails
    {
        public string StatusId { get; set; }
        public string keyField { get; set; }
        public string ErrorDescription { get; set; }
    }
}