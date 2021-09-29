using System;

namespace Bing.EasyExcel.Attributes.Write.Styles
{
    /// <summary>
    /// 标题行高 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
    public class HeadRowHeightAttribute : ExcelAttribute
    {
        /// <summary>
        /// 行高
        /// </summary>
        /// <value>-1：自动设置高度</value>
        public short Height { get; set; } = -1;

        /// <summary>
        /// 初始化一个<see cref="HeadRowHeightAttribute"/>类型的实例
        /// </summary>
        public HeadRowHeightAttribute() { }

        /// <summary>
        /// 初始化一个<see cref="HeadRowHeightAttribute"/>类型的实例
        /// </summary>
        /// <param name="height">行高</param>
        public HeadRowHeightAttribute(short height) => Height = height;
    }
}
