using System;

namespace Bing.EasyExcel.Attributes.Format
{
    /// <summary>
    /// 数值格式化特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class NumberFormatAttribute : ExcelAttribute
    {
        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 舍入方式
        /// </summary>
        public MidpointRounding MidpointRounding { get; set; } = MidpointRounding.ToEven;
    }
}
