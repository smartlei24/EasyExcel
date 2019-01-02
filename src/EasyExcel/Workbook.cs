using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;

namespace EasyExcel {
public class Workbook
    {
        private IWorkbook _workbook;
        private List<dynamic> _sheetBooks = new List<dynamic>();

        public string MimeType { get; }
        public string ExtensionName { get; }

        public Workbook(ExcelType type = ExcelType.XSSF)
        {
            switch (type)
            {
                case ExcelType.HSSF:
                    _workbook = new HSSFWorkbook();
                    break;
                case ExcelType.XSSF:
                    _workbook = new XSSFWorkbook();
                    break;
                //case ExcelType.SXSSF:
                //    _workbook = new SXSSFWorkbook();
                //    break;
            }

            if (type == ExcelType.HSSF)
            {
                MimeType = ExcelMimeType.XLS;
                ExtensionName = ExcelExtensionName.XLS;
            }
            else
            {
                MimeType = ExcelMimeType.XLSX;
                ExtensionName = ExcelExtensionName.XLSX;
            }
        }

        public Sheet<T> AddSheet<T>(string name) where T : class
        {
            var sheet = new Sheet<T>(_workbook, name);
            _sheetBooks.Add(sheet);
            return sheet;
        }

        public IWorkbook Build()
        {
            _sheetBooks.ForEach(i => {
                i.Build();
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
        /// XSSF 的低内存占用版本，可解决其他两者数据量超出65536条后内存溢出的问题，后缀名为 .xlsx, 但目前样式有问题
        /// </summary>
        SXSSF = 2
    }
}
