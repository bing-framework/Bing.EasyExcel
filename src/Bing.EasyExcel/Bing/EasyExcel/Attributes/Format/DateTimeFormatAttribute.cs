using System;

namespace Bing.EasyExcel.Attributes.Format
{
    /// <summary>
    /// 日期时间格式化特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class DateTimeFormatAttribute : ExcelAttribute
    {
        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }
    }
}
