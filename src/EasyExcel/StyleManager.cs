using NPOI.SS.UserModel;
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

        public StyleManager(IWorkbook workbook)
        {
            _workbook = workbook;
        }

        public ICellStyle GetDefaultTitleStyle()
        {
            if (_defaultTitleStyle != null)
            {
                return _defaultTitleStyle;
            }

            var font  = _workbook.CreateFont();
            font.Color = White.Index;
            font.IsBold = true;

            _defaultTitleStyle = _workbook.CreateCellStyle();
            _defaultTitleStyle.Alignment = HorizontalAlignment.Center;
            _defaultTitleStyle.VerticalAlignment = VerticalAlignment.Center;
            _defaultTitleStyle.FillForegroundColor = DarkTeal.Index;
            _defaultTitleStyle.FillPattern = FillPattern.SolidForeground;
            _defaultTitleStyle.SetFont(font);
            return _defaultTitleStyle;
        }

        public ICellStyle GetColumnStyle<T>(Column<T> column)
        {
            if (column.CellStyle != null)
            {
                return column.CellStyle;
            }

            if (column.createStyleFunction != null)
            {
                column.CellStyle = _workbook.CreateCellStyle();
                column.createStyleFunction(column.CellStyle);
                return column.CellStyle;
            }

            var formatString = column.FormatString;
            if (formatString == null || formatString == string.Empty)
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

        private ICellStyle CreateNewBodyStyle(string formatString, Type type)
        {
            var cellStyle = _workbook.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            IDataFormat format = _workbook.CreateDataFormat();
            cellStyle.DataFormat = format.GetFormat(formatString);
            if (IsNumbericType(type))
            {
                cellStyle.Alignment = HorizontalAlignment.Right;
            }
            else
            {
                cellStyle.Alignment = HorizontalAlignment.Left;
            }

            return cellStyle;
        }

        private bool IsNumbericType(Type type)
        {
            return type == typeof(int)
            || type == typeof(short)
            || type == typeof(long)
            || type == typeof(uint)
            || type == typeof(float)
            || type == typeof(double)
            || type == typeof(decimal);
        }

        public ICellStyle GetStripeStyle(ICellStyle baseStyle)
        {
            if (_stripeStyleDic.TryGetValue(baseStyle.Index, out var stripedStyle))
            {
                return stripedStyle;
            }
            stripedStyle = _workbook.CreateCellStyle();
            stripedStyle.CloneStyleFrom(baseStyle);
            stripedStyle.FillForegroundColor = LightCornflowerBlue.Index;
            stripedStyle.FillPattern = FillPattern.SolidForeground;

            _stripeStyleDic.Add(baseStyle.Index, stripedStyle);
            return stripedStyle;
        }

        private Dictionary<Type, string> _defaultFormmat = new Dictionary<Type, string>()
        {
            { typeof(DateTime), "MM/dd/yyyy"},
            { typeof(string), "TEXT"},
            { typeof(char), "TEXT" },
            { typeof(decimal), @"$#,##0.00" },
            { typeof(int), "0" },
            { typeof(short), "0"},
            { typeof(long), "0"},
            { typeof(uint), "0"},
            { typeof(float), "0.00"},
            { typeof(double), "0.00"}
        };
    }
}