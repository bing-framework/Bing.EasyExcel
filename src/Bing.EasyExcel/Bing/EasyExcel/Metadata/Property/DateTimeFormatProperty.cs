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
    }
}
