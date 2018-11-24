using EasyExcel.Extension;
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

        /// <summary>
        /// 列宽度值, NPOI 中使用标准字体个数的宽度来描述宽度
        /// https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Sheet.html#setColumnWidth-int-int-
        /// </summary>
        public float Width { get; private set; } = -1;

        /// <summary>
        /// 用于保存输入值的计算表达式
        /// Expression<Func<T, TResult>>, 因TResult类型不能确定, 所以使用dynamic避免装拆箱
        /// </summary>
        private dynamic valueExpression;
        private dynamic value;
        public Type ValueType { get; private set; }

        public Column(string name)
        {
            Name = name;
        }

        /// <summary>
        /// 指定该列为固定值
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public Column<T> Value<TResult>(TResult value)
        {
            this.value = value;
            ValueType = GetValueType(value.GetType());

            return this;
        }

        /// <summary>
        /// 指定该列为计算值
        /// </summary>
        /// <param name="func"></param>
        /// <returns></returns>
        public Column<T> Value<TResult>(Expression<Func<T, TResult>> func)
        {
            valueExpression = func;
            ValueType = GetValueType(func.ReturnType);

            return this;
        }

        /// <summary>
        /// 指定该列的宽度
        /// </summary>
        /// <param name="width">宽度（以字符宽度为单位）</param>
        /// <returns></returns>
        public Column<T> HasWidth(float width)
        {
            Width = width;
            return this;
        }

        /// <summary>
        /// 指定该列除第一行外的单元格样式
        /// </summary>
        /// <param name="action">设定样式的lamda</param>
        public Column<T> HasStyle(Action<ICellStyle> action)
        {
            setStyleAction = action;
            return this;
        }

        /// <summary>
        /// 该列的位置，从0开始，不指定时按照声明顺序排序
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public Column<T> HasIndex(uint index)
        {
            Index = Convert.ToInt32(index);
            return this;
        }

        /// <summary>
        /// 指定该列的类型
        /// </summary>
        /// <param name="cellType"></param>
        /// <returns></returns>
        public Column<T> HasCellType(CellType cellType = CellType.String)
        {
            CellType = cellType;
            return this;
        }

        /// <summary>
        /// 指定该列的格式化字符串
        /// </summary>
        /// <param name="formatString"></param>
        /// <returns></returns>
        public Column<T> HasFormat(string formatString)
        {
            FormatString = formatString;
            return this;
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

            if (ValueType == typeof(DateTime))
            {
                // 当传入可空DateTime时，当值为 null，应当显示空白，否则转换为 DateTime 显示
                if (cellValue == null)
                {
                    bodyCell.SetCellValue("");
                }
                else
                {
                    bodyCell.SetCellValue(Convert.ToDateTime(cellValue));
                }
                return;
            }

            if (ValueType.IsNumbericType())
            {
                // 传入可空数值类型，当值为null，应当显示 0， 因此全部转换
                bodyCell.SetCellValue(Convert.ToDouble(cellValue));
                return;
            }

            bodyCell.SetCellValue(cellValue == null ? "" : cellValue.ToString());
        }

        private Type GetValueType(Type type)
        {
            // 若为可空类型， 则取其根类型， 如 int? 则返回 Int32
            if (type.IsNullable())
            {
                return type.GetGenericArguments()[0];
            }
            return type;
        }
    }
}
