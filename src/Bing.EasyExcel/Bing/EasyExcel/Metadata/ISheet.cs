using System.Collections.Generic;

namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 工作表
    /// </summary>
    public interface ISheet : IEnumerable<IRow>, IAdapter
    {
        /// <summary>
        /// 工作表名称
        /// </summary>
        string SheetName { get; }

        /// <summary>
        /// 获取单元行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        IRow this[int rowIndex] { get; }

        /// <summary>
        /// 复制单元行
        /// </summary>
        /// <param name="start">开始行索引</param>
        /// <param name="end">结束行索引</param>
        int CopyRows(int start, int end);

        /// <summary>
        /// 移除单元行
        /// </summary>
        /// <param name="start">开始行索引</param>
        /// <param name="end">结束行索引</param>
        int RemoveRows(int start, int end);
    }
}
