using Bing.EasyExcel.Attributes.Format;

namespace Bing.EasyExcel.Metadata.Property
{
    /// <summary>
    /// 日期时间格式化属性
    /// </summary>
    public class DateTimeFormatProperty
    {
        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 是否使用1904年日期窗口的Excel日期
        /// </summary>
        public bool Use1904Windowing { get; set; }

        /// <summary>
        /// 初始化一个<see cref="DateTimeFormatProperty"/>类型的实例
        /// </summary>
        /// <param name="format">日期格式化</param>
        /// <param name="use1904Windowing">是否使用1904年日期窗口的Excel日期</param>
        public DateTimeFormatProperty(string format, bool use1904Windowing)
        {
            Format = format;
            Use1904Windowing = use1904Windowing;
        }

        /// <summary>
        /// 构建
        /// </summary>
        /// <param name="dateTimeFormat">日期时间格式化</param>
        public static DateTimeFormatProperty Build(DateTimeFormatAttribute dateTimeFormat)
        {
            if (dateTimeFormat == null)
                return null;
            return new DateTimeFormatProperty(dateTimeFormat.Format, dateTimeFormat.Use1904Windowing);
        }
    }
}
