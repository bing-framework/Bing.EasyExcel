using System;

namespace Bing.EasyExcel.Attributes.Write.Styles
{
    /// <summary>
    /// 表头样式特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    [Serializable]
    public class HeaderStyleAttribute : StyleAttribute
    {
        /// <summary>
        /// 列宽
        /// </summary>
        public ushort ColumnWidth { get; set; } = 6;
    }
}
