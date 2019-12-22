namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 单元格跨度信息
    /// </summary>
    public struct CellSpan
    {
        /// <summary>
        /// 行跨度
        /// </summary>
        public int RowSpan { get; set; }

        /// <summary>
        /// 列跨度
        /// </summary>
        public int ColSpan { get; set; }

        /// <summary>
        /// 初始化一个<see cref="CellSpan"/>类型的实例
        /// </summary>
        /// <param name="rowSpan">行跨度</param>
        /// <param name="colSpan">列跨度</param>
        public CellSpan(int rowSpan, int colSpan)
        {
            RowSpan = rowSpan;
            ColSpan = colSpan;
        }
    }
}
