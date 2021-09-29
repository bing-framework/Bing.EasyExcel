using System;

namespace Bing.EasyExcel.Attributes.Format
{
    /// <summary>
    /// 日期时间格式化 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false, Inherited = false)]
    public class DateTimeFormatAttribute : ExcelAttribute
    {
        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 是否使用1904年日期窗口的Excel日期
        /// </summary>
        /// <value>
        /// true：使用1904年日期。<br />
        /// false：使用1900年日期。
        /// </value>
        public bool Use1904Windowing { get; set; }

        /// <summary>
        /// 初始化一个<see cref="DateTimeFormatAttribute"/>类型的实例
        /// </summary>
        /// <param name="format">格式化</param>
        public DateTimeFormatAttribute(string format) => Format = format;
    }
}
