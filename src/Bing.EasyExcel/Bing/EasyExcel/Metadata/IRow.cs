using System.Collections.Generic;

namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 单元行
    /// </summary>
    public interface IRow : IEnumerable<ICell>, IAdapter
    {
        /// <summary>
        /// 获取单元格
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        ICell this[int columnIndex] { get; }
    }
}
