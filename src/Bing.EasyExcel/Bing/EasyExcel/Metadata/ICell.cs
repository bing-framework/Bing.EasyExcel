namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 单元格
    /// </summary>
    public interface ICell : IAdapter
    {
        /// <summary>
        /// 行索引
        /// </summary>
        int RowIndex { get; }

        /// <summary>
        /// 列索引
        /// </summary>
        int ColumnIndex { get; }

        /// <summary>
        /// 值
        /// </summary>
        object Value { get; set; }
    }
}
