using System;

namespace Bing.EasyExcel.Attributes.Write.Styles
{
    /// <summary>
    /// 列宽 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field | AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
    public class ColumnWidthAttribute : ExcelAttribute
    {
        /// <summary>
        /// 列宽
        /// </summary>
        /// <value>-1：使用默认列宽</value>
        public int Width { get; set; } = -1;

        /// <summary>
        /// 初始化一个<see cref="ColumnWidthAttribute"/>类型的实例
        /// </summary>
        public ColumnWidthAttribute() { }

        /// <summary>
        /// 初始化一个<see cref="ColumnWidthAttribute"/>类型的实例
        /// </summary>
        /// <param name="width">列宽</param>
        public ColumnWidthAttribute(int width) => Width = width;
    }
}
