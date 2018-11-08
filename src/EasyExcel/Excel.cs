using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;

namespace EasyExcel {
    public class Excel 
    {
        private IWorkbook _workbook;
        private List<dynamic> _sheetBooks = new List<dynamic> ();

        public Excel (ExcelType type) {
            switch (type) {
                case ExcelType.HSSF:
                    _workbook = new HSSFWorkbook ();
                    break;
                case ExcelType.XSSF:
                    _workbook = new XSSFWorkbook ();
                    break;
                case ExcelType.SXSSF:
                    _workbook = new SXSSFWorkbook ();
                    break;
            }
        }

        public SheetBook<T> AddSheetbook<T> (string name, List<T> data) where T : class {
            var sheet = new SheetBook<T> (_workbook, name, data);
            _sheetBooks.Add (sheet);
            return sheet;
        }

        public IWorkbook Build () {
            _sheetBooks.ForEach (i => {
                i.Build ();
            });

            return _workbook;
        }
    }

    public enum ExcelType {
        /// <summary>
        /// 适用于 Excel2003 以前的版本，后缀名为 .xls
        /// </summary>
        HSSF = 0,
        /// <summary>
        /// 适用于 Excel 2007 之后的版本，后缀名为 .xlsx
        /// </summary>
        XSSF = 1,
        /// <summary>
        /// XSSF 的低内存占用版本，可解决其他两者数据量超出65536条后内存溢出的问题
        /// </summary>
        SXSSF = 2
    }
}
