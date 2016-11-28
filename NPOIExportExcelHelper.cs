using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace NPOIExcel.WEB.Helpers
{
    /// <summary>
    /// NPOI Excel Export Helper
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class NPOIExportExcelHelper<T> where T : class
    {
        /// <summary>
        /// 資料表名稱
        /// </summary>
        private string _sheetName = "Sheet 1";
        /// <summary>
        /// 是否為機密文件(Default true)，
        /// 若為True則會於資料表首列加入機密等級欄位
        /// </summary>
        private bool _isSecret = true;
        /// <summary>
        /// 建構式
        /// </summary>
        /// <param name="isSecret">機密等級</param>
        public NPOIExportExcelHelper(bool isSecret = true)
        {
            _isSecret = isSecret;
        }
        /// <summary>
        /// 建構式
        /// </summary>
        /// <param name="sheetName">資料表名稱</param>
        /// <param name="isSecret">機密等級</param>
        public NPOIExportExcelHelper(string sheetName, bool isSecret = true)
        {
            _sheetName = sheetName;
            _isSecret = isSecret;
        }
        /// <summary>
        /// 輸出Excel的檔案串流
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public byte[] ExportToExcel(IEnumerable<T> list)
        {
            var wb = new HSSFWorkbook();    // 建立新的Excel檔
            wb.CreateSheet(_sheetName);     // 建立Sheet
            var st = wb.GetSheet(_sheetName);

            #region 設定標題
            var srcType = list.GetType().UnderlyingSystemType.GetProperty("Item").PropertyType; // Model型別
            var srcProps = srcType.GetProperties(); // Model內容
            var rowTitle = st.CreateRow(0); // 建立Title列
            var defaultCol = 0; // Model沒有設定Column時以此為主
            foreach (var prop in srcProps)
            {
                var attr = prop.GetCustomAttributes(typeof(ExcelExportAttribute), true).FirstOrDefault();
                if (attr != null)   // 判斷該Prop是否有套用ExcelExport的Attribute，若有才將該欄位輸出至Excel
                {
                    var title = ((ExcelExportAttribute)attr).Title;
                    var column = ((ExcelExportAttribute)attr).Column == -1 ? defaultCol : ((ExcelExportAttribute)attr).Column;
                    rowTitle.CreateCell(column).SetCellValue(title);
                    defaultCol++;
                }
            }
            #endregion

            #region 設定內容
            if (list.Count() > 0)
            {
                foreach(var item in list)
                {
                    var props = item.GetType().GetProperties();
                    var row = st.CreateRow(st.LastRowNum + 1);   // index 0 is title
                    defaultCol = 0;
                    foreach (var prop in props)
                    {
                        var attr = prop.GetCustomAttributes(typeof(ExcelExportAttribute), true).FirstOrDefault();
                        if (attr != null)
                        {
                            var type = prop.PropertyType.FullName;
                            var value = prop.GetValue(item);
                            var column = ((ExcelExportAttribute)attr).Column == -1 ? defaultCol : ((ExcelExportAttribute)attr).Column;

                            #region 依照型別設定資料
                            if (type.Equals(typeof(string).FullName))
                            {
                                // 字串型別
                                row.CreateCell(column).SetCellValue((string)value);
                            }
                            else if (type.Equals(typeof(DateTime).FullName))
                            {
                                // 日期型別
                                row.CreateCell(column).SetCellValue((DateTime)value);
                            }
                            else if (type.Equals(typeof(bool).FullName))
                            {
                                // 布林型別
                                row.CreateCell(column).SetCellValue((bool)value);
                            }
                            else
                            {
                                // 其他型別
                                row.CreateCell(column).SetCellValue(value.ToString());
                            }
                            #endregion

                            defaultCol++;
                        }
                    }
                }
            }
            #endregion

            #region 自動調整欄寬
            for (var i = 0; i < st.GetRow(0).LastCellNum; i++)
            {
                st.AutoSizeColumn(i);
            }
            #endregion

            #region 機密等級
            if (_isSecret)
            {
                // 所有資料往下移一列
                st.ShiftRows(0, st.LastRowNum, 1);
                // 合併列
                st.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, st.GetRow(1).LastCellNum - 1));
                st.CreateRow(0).CreateCell(0).SetCellValue("機密等級：");
            }
            #endregion

            MemoryStream ms = new MemoryStream();
            wb.Write(ms);

            return ms.ToArray();
        }
    }

    /// <summary>
    /// Model Attribute
    /// </summary>
    public class ExcelExportAttribute : Attribute
    {
        /// <summary>
        /// 欄位(Default)
        /// 預設為-1，沒有傳入時會依照Model內的排序來安排欄位
        /// </summary>
        private int _Column = -1;
        /// <summary>
        /// 標題
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// 欄位
        /// </summary>
        public int Column { get; set; }
    }
}