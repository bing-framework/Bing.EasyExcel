using Bing.EasyExcel.Npoi.Metadata;

namespace Bing.EasyExcel.Npoi.Extensions
{
    /// <summary>
    /// NPOI单元行(<see cref="NPOI.SS.UserModel.IRow"/>) 扩展
    /// </summary>
    public static class RowExtensions
    {
        /// <summary>
        /// 获取适配器
        /// </summary>
        /// <param name="row">NPOI单元行</param>
        internal static Row GetAdapter(this NPOI.SS.UserModel.IRow row) => null == row ? null : new Row(row);
    }
}
