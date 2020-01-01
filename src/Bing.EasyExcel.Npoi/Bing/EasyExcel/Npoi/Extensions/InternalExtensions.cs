using Bing.EasyExcel.Npoi.Metadata;

namespace Bing.EasyExcel.Npoi.Extensions
{
    /// <summary>
    /// 内部扩展
    /// </summary>
    internal static class InternalExtensions
    {
        /// <summary>
        /// 获取适配器
        /// </summary>
        /// <param name="cell">NPOI单元格</param>
        internal static Cell GetAdapter(this NPOI.SS.UserModel.ICell cell) => null == cell ? null : new Cell(cell);

        /// <summary>
        /// 获取适配器
        /// </summary>
        /// <param name="row">NPOI单元行</param>
        internal static Row GetAdapter(this NPOI.SS.UserModel.IRow row) => null == row ? null : new Row(row);

        /// <summary>
        /// 获取适配器
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        internal static Sheet GetAdapter(this NPOI.SS.UserModel.ISheet sheet) => null == sheet ? null : new Sheet(sheet);

        /// <summary>
        /// 获取适配器
        /// </summary>
        /// <param name="workbook">NPOI工作簿</param>
        internal static Workbook GetAdapter(this NPOI.SS.UserModel.IWorkbook workbook) =>
            null == workbook ? null : new Workbook(workbook);
    }
}
