using NPOIExcel.WEB.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace NPOIExcel.WEB.Models
{
    public class ExportVM
    {
        /// <summary>
        /// 營業點代碼
        /// </summary>
        [ExcelExport(Title = "營業點代碼", Column = 0)]
        public string BC_CDE { get; set; }
        /// <summary>
        /// 營業店點
        /// </summary>
        [ExcelExport(Title = "營業店點", Column = 1)]
        public string BC_NAME { get; set; }
    }
}