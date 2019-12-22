namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 单元格范围
    /// </summary>
    public class CellRange
    {
        /// <summary>
        /// 起始行索引
        /// </summary>
        public int FirstRow { get; set; }

        /// <summary>
        /// 结束行索引
        /// </summary>
        public int LastRow { get; set; }

        /// <summary>
        /// 起始列索引
        /// </summary>
        public int FirstCol { get; set; }

        /// <summary>
        /// 结束列索引
        /// </summary>
        public int LastCol { get; set; }

        /// <summary>
        /// 初始化一个<see cref="CellRange"/>类型的实例
        /// </summary>
        /// <param name="firstRow">起始行索引</param>
        /// <param name="lastRow">结束行索引</param>
        /// <param name="firstCol">起始列索引</param>
        /// <param name="lastCol">结束列索引</param>
        public CellRange(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstCol = firstCol;
            LastCol = lastCol;
        }
    }
}
