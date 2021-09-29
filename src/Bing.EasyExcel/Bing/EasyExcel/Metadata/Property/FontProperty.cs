using Bing.EasyExcel.Attributes.Write.Styles;

namespace Bing.EasyExcel.Metadata.Property
{
    /// <summary>
    /// 字体属性
    /// </summary>
    public class FontProperty
    {
        /// <summary>
        /// 字体
        /// </summary>
        public string FontFamily { get; set; }

        /// <summary>
        /// 字体大小
        /// </summary>
        public short FontSize { get; set; }

        /// <summary>
        /// 斜体。是否使用斜体
        /// </summary>
        public bool Italic { get; set; }

        /// <summary>
        /// 删除线。是否在文本中使用删除线
        /// </summary>
        public bool Strikeout { get; set; }

        /// <summary>
        /// 字体颜色
        /// </summary>
        public short Color { get; set; }

        /// <summary>
        /// 字体偏移。正常、上标、下标
        /// </summary>
        /// <value>
        /// 0：正常<br />
        /// 1：上标<br />
        /// 2：下标<br />
        /// </value>
        public short TypeOffset { get; set; }

        /// <summary>
        /// 文本下划线类型
        /// </summary>
        /// <value>
        /// 0：正常<br />
        /// 1：下划线<br />
        /// 2：双下划线<br />
        /// 3：强调线<br />
        /// 4：双强调线<br />
        /// </value>
        public byte Underline { get; set; }

        /// <summary>
        /// 加粗
        /// </summary>
        public bool Bold { get; set; }

        /// <summary>
        /// 构建
        /// </summary>
        /// <param name="headFontStyle">标题字体样式</param>
        public static FontProperty Build(HeadFontStyleAttribute headFontStyle)
        {
            if (headFontStyle == null)
                return null;
            var styleProperty = new FontProperty();
            if (string.IsNullOrWhiteSpace(headFontStyle.FontFamily))
                styleProperty.FontFamily = headFontStyle.FontFamily;
            if (headFontStyle.FontSize >= 0)
                styleProperty.FontSize = headFontStyle.FontSize;
            styleProperty.Italic = headFontStyle.Italic;
            styleProperty.Strikeout = headFontStyle.Strikeout;
            if (headFontStyle.Color >= 0)
                styleProperty.Color = headFontStyle.Color;
            if (headFontStyle.TypeOffset >= 0)
                styleProperty.TypeOffset = headFontStyle.TypeOffset;
            if (headFontStyle.Underline > 0)
                styleProperty.Underline = headFontStyle.Underline;
            styleProperty.Bold = headFontStyle.Bold;
            return styleProperty;
        }

        /// <summary>
        /// 构建
        /// </summary>
        /// <param name="contentFontStyle">内容字体样式</param>
        public static FontProperty Build(ContentFontStyleAttribute contentFontStyle)
        {
            if (contentFontStyle == null)
                return null;
            var styleProperty = new FontProperty();
            if (string.IsNullOrWhiteSpace(contentFontStyle.FontFamily))
                styleProperty.FontFamily = contentFontStyle.FontFamily;
            if (contentFontStyle.FontSize >= 0)
                styleProperty.FontSize = contentFontStyle.FontSize;
            styleProperty.Italic = contentFontStyle.Italic;
            styleProperty.Strikeout = contentFontStyle.Strikeout;
            if (contentFontStyle.Color >= 0)
                styleProperty.Color = contentFontStyle.Color;
            if (contentFontStyle.TypeOffset >= 0)
                styleProperty.TypeOffset = contentFontStyle.TypeOffset;
            if (contentFontStyle.Underline > 0)
                styleProperty.Underline = contentFontStyle.Underline;
            styleProperty.Bold = contentFontStyle.Bold;
            return styleProperty;

        }
    }
}
