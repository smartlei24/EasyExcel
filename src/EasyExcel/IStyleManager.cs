using NPOI.SS.UserModel;

namespace EasyExcel
{
    internal interface IStyleManager
    {
        ICellStyle GetDefaultTitleStyle();

        ICellStyle GetColumnStyle<T>(Column<T> column);

        ICellStyle GetStripeStyle(ICellStyle baseStyle);
    }
}
