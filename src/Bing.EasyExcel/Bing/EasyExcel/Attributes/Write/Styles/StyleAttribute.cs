using Bing.EasyExcel.Enums.Styles;

namespace Bing.EasyExcel.Attributes.Write.Styles
{
    /// <summary>
    /// 样式特性
    /// </summary>
    public abstract class StyleAttribute : ExcelAttribute
    {
        /// <summary>
        /// 高度
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// 文本颜色
        /// </summary>
        public string TextColor { get; set; }

        /// <summary>
        /// 背景色
        /// </summary>
        public string BackgroundColor { get; set; }

        /// <summary>
        /// 前景色
        /// </summary>
        public string ForegroundColor { get; set; }

        /// <summary>
        /// 文字对齐
        /// </summary>
        public HorizontalAlignment TextAlign { get; set; }

        /// <summary>
        /// 垂直对齐
        /// </summary>
        public VerticalAlignment VerticalAlign { get; set; }

        /// <summary>
        /// 填充模式
        /// </summary>
        public FillPattern FillPattern { get; set; }

        /// <summary>
        /// 是否粗体
        /// </summary>
        public short FontWeight { get; set; }

        /// <summary>
        /// 字体
        /// </summary>
        public string FontFamily { get; set; }

        /// <summary>
        /// 字体大小
        /// </summary>
        public short FontSize { get; set; }

        /// <summary>
        /// 是否斜体
        /// </summary>
        public bool IsItalic { get; set; }

        /// <summary>
        /// 初始化一个<see cref="StyleAttribute"/>类型的实例
        /// </summary>
        protected StyleAttribute()
        {
            TextAlign = HorizontalAlignment.Center;
            VerticalAlign = VerticalAlignment.Top;
            FontSize = -1;
            FillPattern = FillPattern.SolidForeground;
        }
    }
}
