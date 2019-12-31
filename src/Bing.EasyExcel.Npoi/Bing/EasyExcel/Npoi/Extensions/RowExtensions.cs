namespace Bing.EasyExcel.Npoi.Extensions
{
    /// <summary>
    /// NPOI单元行(<see cref="NPOI.SS.UserModel.IRow"/>) 扩展
    /// </summary>
    public static class RowExtensions
    {
        /// <summary>
        /// 清空内容
        /// </summary>
        /// <param name="row">NPOI单元行</param>
        public static NPOI.SS.UserModel.IRow ClearContent(this NPOI.SS.UserModel.IRow row)
        {
            foreach (var cell in row.Cells)
                cell.SetCellValue(string.Empty);
            return row;
        }
    }
}
