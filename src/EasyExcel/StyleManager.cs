using EasyExcel.Extension;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using static NPOI.HSSF.Util.HSSFColor;

namespace EasyExcel
{
internal class StyleManager : IStyleManager
    {
        private IWorkbook _workbook;
        private ICellStyle _defaultTitleStyle;
        private Dictionary<string, ICellStyle> _styleDic = new Dictionary<string, ICellStyle>();
        private Dictionary<int, ICellStyle> _stripeStyleDic = new Dictionary<int, ICellStyle>();

        private IColor _stripedCellBgColor;
        private IColor _titleCellBgColor;

        /// <summary>
        /// 用于指向 HSSF 上一个自定义颜色的索引，NPOI 中每个色板最多只能包含 64 中颜色，一般从第九种颜色开始自定义，覆盖原有的该Index的颜色的值。
        /// </summary>
        private short _customColorIndex = 8;

        public StyleManager(IWorkbook workbook)
        {
            _workbook = workbook;
        }

        public ICellStyle GetTitleStyle()
        {
            if (_defaultTitleStyle != null)
            {
                return _defaultTitleStyle;
            }

            var font = _workbook.CreateFont();
            font.Color = White.Index;
            font.IsBold = true;

            _defaultTitleStyle = _workbook.CreateCellStyle();
            _defaultTitleStyle.Alignment = HorizontalAlignment.Center;
            _defaultTitleStyle.VerticalAlignment = VerticalAlignment.Center;
            IColor color = GetWorkbookColor(54, 96, 146);
            SetStyleFillForegroundColor(_defaultTitleStyle, color);
            _defaultTitleStyle.FillPattern = FillPattern.SolidForeground;
            _defaultTitleStyle.SetFont(font);
            return _defaultTitleStyle;
        }

        /// <summary>
        /// 根据列的数据类型来获取默认样式
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public ICellStyle GetBodyCellStyle<T>(Column<T> column)
        {
            var formatString = column.FormatString;
            if (formatString.IsNullOrEmpty())
            {
                formatString = _defaultFormmat.GetValueOrDefault(column.ValueType, "General");
            }

            if (_styleDic.TryGetValue(formatString, out var cellStyle))
            {
                return cellStyle;
            }
            else
            {
                // 若目前还没有该格式的样式， 则创建样式并加入
                cellStyle = CreateNewBodyStyle(formatString, column.ValueType);
            }

            _styleDic.Add(formatString, cellStyle);
            return cellStyle;
        }

        /// <summary>
        /// 获取样式的"条纹"样式
        /// </summary>
        /// <param name="baseStyle"></param>
        /// <returns></returns>
        public ICellStyle GetStripeStyle(ICellStyle baseStyle)
        {
            if (_stripeStyleDic.TryGetValue(baseStyle.Index, out var stripedStyle))
            {
                return stripedStyle;
            }
            stripedStyle = _workbook.CreateCellStyle();
            stripedStyle.CloneStyleFrom(baseStyle);

            var color = _stripedCellBgColor ?? GetWorkbookColor(224, 235, 252);
            SetStyleFillForegroundColor(stripedStyle, color);

            stripedStyle.FillPattern = FillPattern.SolidForeground;

            _stripeStyleDic.Add(baseStyle.Index, stripedStyle);
            return stripedStyle;
        }

        private Dictionary<Type, string> _defaultFormmat = new Dictionary<Type, string>()
        {
            { typeof(DateTime), "MM/dd/yyyy"},
            { typeof(string), "TEXT"},
            { typeof(char), "TEXT" },
            { typeof(decimal), "$#,##0.00" },
            { typeof(int), "0" },
            { typeof(short), "0"},
            { typeof(long), "0"},
            { typeof(uint), "0"},
            { typeof(float), "0.00"},
            { typeof(double), "0.00"}
        };

        private ICellStyle CreateNewBodyStyle(string formatString, Type type)
        {
            var cellStyle = _workbook.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            IDataFormat format = _workbook.CreateDataFormat();
            cellStyle.DataFormat = format.GetFormat(formatString);
            if (type.IsNumbericType())
            {
                cellStyle.Alignment = HorizontalAlignment.Right;
            }
            else
            {
                cellStyle.Alignment = HorizontalAlignment.Left;
            }

            return cellStyle;
        }

        private IColor GetWorkbookColor(byte r, byte g, byte b)
        {
            if (_workbook is HSSFWorkbook)
            {
                var palette = ((HSSFWorkbook)_workbook).GetCustomPalette();
                var color = palette.FindColor(r, g, b);
                if (color == null)
                {
                    palette.SetColorAtIndex(_customColorIndex, r, g, b);
                    color = palette.FindColor(r, g, b);
                    _customColorIndex++;
                }

                return color;
            }
            else
            {
                 return new XSSFColor(new byte[] { r, g, b });
            }
        }

        private void SetStyleFillForegroundColor(ICellStyle style, IColor color)
        {
            if (style is HSSFCellStyle)
            {
                style.FillForegroundColor = color.Indexed;
            }
            else
            {
                ((XSSFCellStyle)style).FillForegroundXSSFColor = color as XSSFColor;
            }
        }
    }
}
