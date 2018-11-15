using NPOI.SS.UserModel;
using System;
using System.Linq.Expressions;

namespace EasyExcel
{
    public class Column<T>
    {
        public string Name { get; private set; }
        public int Index { get; internal set; } = -1;
        public string FormatString { get; private set; }
        public CellType CellType { get; private set; } = CellType.String;
        public ICellStyle CellStyle { get; internal set; }
        internal Action<ICellStyle> setStyleAction;
        private dynamic valueExpression;
        public Type ValueType { get; private set; }
        public int Width { get; private set; } = 9;
        private dynamic value;

        public Column(string name)
        {
            Name = name;
        }

        /// <summary>
        /// 指定该列为固定值
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public Column<T> WithValue<TResult>(TResult value)
        {
            this.value = value;
            ValueType = value.GetType();
            return this;
        }

        /// <summary>
        /// 指定该列为计算值
        /// </summary>
        /// <param name="func"></param>
        /// <returns></returns>
        public Column<T> WithValue<TResult>(Expression<Func<T, TResult>> func)
        {
            valueExpression = func;
            ValueType = func.ReturnType;
            return this;
        }

        /// <summary>
        /// 设置该列的宽度
        /// </summary>
        /// <param name="width"></param>
        /// <returns></returns>
        public Column<T> WithWidth(int width)
        {
            Width = width;
            return this;
        }

        /// <summary>
        /// 设置该列的单元格样式
        /// </summary>
        /// <param name="action"></param>
        public Column<T> WithBodyStyle(Action<ICellStyle> action)
        {
            setStyleAction = action;
            return this;
        }

        /// <summary>
        /// 该列的位置，从0开始，不指定时按照声明顺序排序
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public Column<T> WithIndex(uint index)
        {
            Index = Convert.ToInt32(index);
            return this;
        }

        /// <summary>
        /// 指定该列的类型
        /// </summary>
        /// <param name="cellType"></param>
        /// <returns></returns>
        public Column<T> WithCellType(CellType cellType)
        {
            CellType = cellType;
            return this;
        }

        /// <summary>
        /// 指定该列的格式化字符串
        /// </summary>
        /// <param name="formatString"></param>
        /// <returns></returns>
        public Column<T> WithFormat(string formatString)
        {
            FormatString = formatString;
            return this;
        }

        internal void SetTitleCell(ICell headCell)
        {
            headCell.SetCellValue(Name);
            headCell.SetCellType(CellType.String);
        }

        internal void SetBodyCell(ICell bodyCell, T data)
        {
            bodyCell.SetCellType(CellType);

            var cellValue = value ?? valueExpression.Compile()(data);
            if (ValueType == typeof(bool))
            {
                bodyCell.SetCellValue(Convert.ToBoolean((cellValue)));
                return;
            }

            if (ValueType == typeof(string)|| ValueType == typeof(char))
            {
                bodyCell.SetCellValue(cellValue as string);
                return;
            }

            if (ValueType == typeof(DateTime))
            {
                bodyCell.SetCellValue(Convert.ToDateTime(cellValue));
                return;
            }

            bodyCell.SetCellValue(Convert.ToDouble(cellValue));
        }
    }
}