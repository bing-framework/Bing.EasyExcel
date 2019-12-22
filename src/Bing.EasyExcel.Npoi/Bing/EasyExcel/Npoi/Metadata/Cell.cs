using Bing.EasyExcel.Metadata;

namespace Bing.EasyExcel.Npoi.Metadata
{
    /// <summary>
    /// 单元格
    /// </summary>
    public class Cell : ICell
    {
        /// <summary>
        /// NPOI单元格
        /// </summary>
        public NPOI.SS.UserModel.ICell NpoiCell { get; private set; }

        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex => NpoiCell.RowIndex;

        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex => NpoiCell.ColumnIndex;

        /// <summary>
        /// 值
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// 初始化一个<see cref="Cell"/>类型的实例
        /// </summary>
        /// <param name="npoiCell">NPOI单元格</param>
        public Cell(NPOI.SS.UserModel.ICell npoiCell) => NpoiCell = npoiCell;

        /// <summary>
        /// 获取原始值
        /// </summary>
        public object GetOriginal() => NpoiCell;
    }
}
