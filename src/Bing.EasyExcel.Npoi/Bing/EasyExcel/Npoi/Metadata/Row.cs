using System.Collections;
using System.Collections.Generic;
using Bing.EasyExcel.Metadata;
using Bing.EasyExcel.Npoi.Extensions;

namespace Bing.EasyExcel.Npoi.Metadata
{
    /// <summary>
    /// 单元行
    /// </summary>
    public class Row : IRow
    {
        /// <summary>
        /// NPOI单元行
        /// </summary>
        public NPOI.SS.UserModel.IRow NpoiRow { get; private set; }

        /// <summary>
        /// 获取单元格
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        public ICell this[int columnIndex] => NpoiRow.GetCell(columnIndex).GetAdapter();

        /// <summary>
        /// 初始化一个<see cref="Row"/>类型的实例
        /// </summary>
        /// <param name="npoiRow">NPOI单元行</param>
        public Row(NPOI.SS.UserModel.IRow npoiRow) => NpoiRow = npoiRow;

        /// <summary>
        /// 获取集合
        /// </summary>
        public IEnumerator<ICell> GetEnumerator()
        {
            foreach (NPOI.SS.UserModel.ICell npoiCell in NpoiRow)
                yield return npoiCell.GetAdapter();
        }

        /// <summary>
        /// 获取集合
        /// </summary>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        /// <summary>
        /// 获取原始值
        /// </summary>
        public object GetOriginal() => NpoiRow;
    }
}
