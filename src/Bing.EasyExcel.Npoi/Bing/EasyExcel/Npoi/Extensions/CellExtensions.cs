using Bing.EasyExcel.Npoi.Metadata;

namespace Bing.EasyExcel.Npoi.Extensions
{
    /// <summary>
    /// NPOI单元格(<see cref="NPOI.SS.UserModel.ICell"/>) 扩展
    /// </summary>
    public static class CellExtensions
    {
        /// <summary>
        /// 获取适配器
        /// </summary>
        /// <param name="cell">NPOI单元格</param>
        internal static Cell GetAdapter(this NPOI.SS.UserModel.ICell cell) => null == cell ? null : new Cell(cell);
    }
}
