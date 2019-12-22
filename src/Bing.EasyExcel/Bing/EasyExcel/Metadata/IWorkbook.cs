using System.Collections.Generic;

namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 工作簿
    /// </summary>
    public interface IWorkbook : IEnumerable<ISheet>, IAdapter
    {
        /// <summary>
        /// 获取工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        ISheet this[string sheetName] { get; }

        /// <summary>
        /// 保存到流
        /// </summary>
        byte[] SaveToBuffer();
    }
}
