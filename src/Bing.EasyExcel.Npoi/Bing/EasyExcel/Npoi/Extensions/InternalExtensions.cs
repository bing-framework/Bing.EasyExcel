using Bing.EasyExcel.Npoi.Metadata;
using NPOI.SS.UserModel;

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
        internal static Cell GetAdapter(this ICell cell) => null == cell ? null : new Cell(cell);

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
    }
}
