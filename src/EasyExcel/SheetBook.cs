using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EasyExcel
{
    public class SheetBook<T> where T : class
    {
        private ISheet _sheet;
        private readonly IStyleManager _styleManager;

        public List<T> Data { get; private set; }
        public List<Column<T>> Columns { get; private set; } = new List<Column<T>>();

        public float TitleRowHight { get; private set; } = 30;
        public float BodyRowHight { get; private set; } = 20;

        public ICellStyle TitleStyle { get; private set; }

        public bool FreezeTitle { get; private set; } = true;
        public bool OutBorder { get; private set; } = true;
        public bool HasStriped { get; private set; } = true;

        internal SheetBook(IWorkbook workbook, string name, List<T> data)
        {
            _sheet = workbook.CreateSheet(name);
            _styleManager = new StyleManager(_sheet.Workbook);
            Data = data;
        }

        /// <summary>
        /// 增加一列
        /// </summary>
        /// <param name="columnName">显示的列名</param>
        /// <returns></returns>
        public Column<T> HasColumn(string columnName)
        {
            var column = new Column<T>(columnName);
            Columns.Add(column);

            return column;
        }

        /// <summary>
        /// 是否固定表头
        /// </summary>
        /// <param name="freeze"></param>
        /// <returns></returns>
        public SheetBook<T> HasFreezeTitle(bool freeze = true)
        {
            FreezeTitle = freeze;
            return this;
        }

        /// <summary>
        /// 是否含有外边框
        /// </summary>
        /// <param name="freeze"></param>
        /// <returns></returns>
        public SheetBook<T> HasOutBolder(bool outBolder = true)
        {
            OutBorder = outBolder;
            return this;
        }

        /// <summary>
        /// 设置标题的单元格样式
        /// </summary>
        /// <param name="action"></param>
        public SheetBook<T> WithTitleStyle(Action<ICellStyle> action)
        {
            TitleStyle = _sheet.Workbook.CreateCellStyle();
            action(TitleStyle);
            return this;
        }

        /// <summary>
        /// 设置标题行的行高
        /// </summary>
        /// <param name="height"></param>
        public SheetBook<T> WithTitleRowHeight(float height = 30f)
        {
            TitleRowHight = height;
            return this;
        }

        /// <summary>
        /// 设置内容的行高
        /// </summary>
        /// <param name="height"></param>
        /// <returns></returns>
        public SheetBook<T> WithBodyRowHeight(float height = 20f)
        {
            BodyRowHight = height;
            return this;
        }

        /// <summary>
        /// 设置内容是否含有行间隔条纹
        /// </summary>
        /// <param name="has"></param>
        /// <returns></returns>
        public SheetBook<T> HasStripedRow(bool has = true)
        {
            HasStriped = has;
            return this;
        }

        /// <summary>
        /// 生成表格
        /// </summary>
        /// <returns></returns>
        public ISheet Build()
        {
            SoftColumms();
            BuildTitle();
            BuildBody();
            BuildStyle();
            return _sheet;
        }

        private void BuildTitle()
        {
            var row = _sheet.CreateRow(0);
            ICell cell;

            row.HeightInPoints = TitleRowHight;

            for (int i = 0; i < Columns.Count; i++)
            {
                cell = row.CreateCell(i, CellType.String);
                Columns[i].SetTitleCell(cell);
            }

            if (FreezeTitle)
            {
                _sheet.CreateFreezePane(Columns.Count, 1);
            }
        }

        private void BuildBody()
        {
            IRow row;
            ICell cell;
            int rowIndex;

            for (int i = 0; i < Data.Count; i++)
            {
                rowIndex = i + 1;
                row = _sheet.CreateRow(rowIndex);
                row.HeightInPoints = BodyRowHight;
                for (int colIndex = 0; colIndex < Columns.Count; colIndex++)
                {
                    cell = row.CreateCell(colIndex);
                    Columns[colIndex].SetBodyCell(cell, Data[i]);
                }
            }
        }

        private void BuildStyle()
        {
            BuildColumnWidth();
            BuildCellStyle();
            BuildOutBorder();
        }

        private void BuildColumnWidth()
        {
            for (int i = 0; i < Columns.Count; i++)
            {

                _sheet.SetColumnWidth(i, Columns[i].Width * 256);
            }
        }

        private void BuildCellStyle()
        {
            for (int i = 0; i < Columns.Count; i++)
            {
                TitleStyle = TitleStyle ?? _styleManager.GetDefaultTitleStyle();
                _sheet.GetRow(0).GetCell(i).CellStyle = TitleStyle;

                ICellStyle style;
                if (Columns[i].setStyleAction != null)
                {
                    style = _sheet.Workbook.CreateCellStyle();
                    Columns[i].setStyleAction(style);
                }
                else
                {
                    style = _styleManager.GetColumnStyle<T>(Columns[i]);
                }

                var stripedStyle = _styleManager.GetStripeStyle(style);
                for (int r = 1; r <= Data.Count; r++)
                {
                    _sheet.GetRow(r).GetCell(i).CellStyle = HasStriped && r % 2 == 0 ? stripedStyle : style;
                }
            }
        }

        private void BuildOutBorder()
        {
            if (OutBorder)
            {
                var maxColIndex = Columns.Count - 1;
                var maxRowIndex = Data.Count;

                // 处理四个边角
                ICellStyle leftTopStyle = _sheet.Workbook.CreateCellStyle();
                ICell leftTopCell = _sheet.GetRow(0).GetCell(0);
                leftTopStyle.CloneStyleFrom(leftTopCell.CellStyle);
                leftTopStyle.BorderLeft = BorderStyle.Medium;
                leftTopStyle.BorderTop = BorderStyle.Medium;
                leftTopStyle.BorderRight = BorderStyle.Thin;
                leftTopStyle.BorderBottom = BorderStyle.Thin;
                leftTopCell.CellStyle = leftTopStyle;

                ICellStyle leftTailStyle = _sheet.Workbook.CreateCellStyle();
                ICell leftTailCell = _sheet.GetRow(maxRowIndex).GetCell(0);
                leftTailStyle.CloneStyleFrom(leftTailCell.CellStyle);
                leftTailStyle.BorderLeft = BorderStyle.Medium;
                leftTailStyle.BorderBottom = BorderStyle.Medium;
                leftTailCell.CellStyle = leftTailStyle;

                ICellStyle rightTopStyle = _sheet.Workbook.CreateCellStyle();
                ICell rightTopCell = _sheet.GetRow(0).GetCell(maxColIndex);
                rightTopStyle.CloneStyleFrom(rightTopCell.CellStyle);
                rightTopStyle.BorderLeft = BorderStyle.Thin;
                rightTopStyle.BorderTop = BorderStyle.Medium;
                leftTopStyle.BorderRight = BorderStyle.Medium;
                leftTopStyle.BorderBottom = BorderStyle.Thin;
                rightTopCell.CellStyle = rightTopStyle;

                ICellStyle rightTailStyle = _sheet.Workbook.CreateCellStyle();
                ICell rightTailCell = _sheet.GetRow(maxRowIndex).GetCell(maxColIndex);
                rightTailStyle.CloneStyleFrom(rightTailCell.CellStyle);
                rightTailStyle.BorderRight = BorderStyle.Medium;
                rightTailStyle.BorderBottom = BorderStyle.Medium;
                rightTailCell.CellStyle = rightTailStyle;

                // 处理除四个角外的表头的边框
                TitleStyle.BorderTop = BorderStyle.Medium;
                TitleStyle.BorderBottom = BorderStyle.Thin;
                TitleStyle.BorderRight = BorderStyle.Thin;

                // 处理除边角外第一列的左边框
                ICellStyle fristColStyle = _sheet.Workbook.CreateCellStyle();
                fristColStyle.CloneStyleFrom(_styleManager.GetColumnStyle(Columns[0]));
                fristColStyle.BorderLeft = BorderStyle.Medium;
                var fristColStripedStyle = _styleManager.GetStripeStyle(fristColStyle);
                fristColStripedStyle.BorderLeft = BorderStyle.Medium;
                for (int i = 1; i < maxRowIndex; i++)
                {
                    _sheet.GetRow(i).GetCell(0).CellStyle = HasStriped && i % 2 == 0 ? fristColStripedStyle : fristColStyle;
                }

                // 处理除边角外最后一列的右边框
                ICellStyle lastColStyle = _sheet.Workbook.CreateCellStyle();
                lastColStyle.CloneStyleFrom(_styleManager.GetColumnStyle(Columns[maxColIndex]));
                lastColStyle.BorderRight = BorderStyle.Medium;
                var lastColStripedStyle = _styleManager.GetStripeStyle(lastColStyle);
                lastColStripedStyle.BorderRight = BorderStyle.Medium;
                for (int i = 1; i < maxRowIndex; i++)
                {
                    _sheet.GetRow(i).GetCell(maxColIndex).CellStyle = HasStriped && i % 2 == 0 ? lastColStripedStyle : lastColStyle;
                }

                // 处理除边角外最后一行的下边框
                var lastRow = _sheet.GetRow(maxRowIndex);
                for (int i = 1; i < maxColIndex; i++)
                {
                    var cell = lastRow.GetCell(i);
                    var style = _sheet.Workbook.CreateCellStyle();
                    style.CloneStyleFrom(cell.CellStyle);
                    style.BorderBottom = BorderStyle.Medium;
                    cell.CellStyle = style;
                }
            }
        }

        /// <summary>
        /// 设置列的顺序
        /// </summary>
        private void SoftColumms()
        {
            var autoSortColumns = Columns.Where(i => i.Index < 0).ToList();
            Columns.Where(i => i.Index >= 0)
                .OrderBy(i => i.Index)
                .ToList()
                .ForEach(i => autoSortColumns.Insert(i.Index, i));
            Columns = autoSortColumns;
        }
    }
}
