using NPOI.SS.UserModel;

namespace EasyExcel
{
    internal interface IStyleManager
    {
        ICellStyle GetTitleStyle();

        ICellStyle GetBodyCellStyle<T>(Column<T> column);

        ICellStyle GetStripeStyle(ICellStyle baseStyle);
    }
}
