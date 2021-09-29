namespace Bing.EasyExcel.Write.Metadata.Styles
{
    /// <summary>
    /// 写入字体
    /// </summary>
    public class WriteFont
    {
        /// <summary>
        /// 字体
        /// </summary>
        public string FontFamily { get; set; }

        /// <summary>
        /// 字体大小
        /// </summary>
        public short? FontSize { get; set; }

        /// <summary>
        /// 斜体。是否使用斜体
        /// </summary>
        public bool? Italic { get; set; }

        /// <summary>
        /// 删除线。是否在文本中使用删除线
        /// </summary>
        public bool? Strikeout { get; set; }

        /// <summary>
        /// 字体颜色
        /// </summary>
        public short? Color { get; set; }

        /// <summary>
        /// 字体偏移。正常、上标、下标
        /// </summary>
        /// <value>
        /// 0：正常<br />
        /// 1：上标<br />
        /// 2：下标<br />
        /// </value>
        public short? TypeOffset { get; set; }

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
        public byte? Underline { get; set; }

        /// <summary>
        /// 加粗
        /// </summary>
        public bool? Bold { get; set; }

        /// <summary>
        /// 合并
        /// </summary>
        /// <param name="source">源数据</param>
        /// <param name="target">目标数据</param>
        public static void Merge(WriteFont source, WriteFont target)
        {
            if (source == null || target == null)
                return;
            if (string.IsNullOrWhiteSpace(source.FontFamily))
                target.FontFamily = source.FontFamily;
            if (source.FontSize != null)
                target.FontSize = source.FontSize;
            if (source.Italic != null)
                target.Italic = source.Italic;
            if (source.Strikeout != null)
                target.Strikeout = source.Strikeout;
            if (source.Color != null)
                target.Color = source.Color;
            if (source.TypeOffset != null)
                target.TypeOffset = source.TypeOffset;
            if (source.Underline != null)
                target.Underline = source.Underline;
            if (source.Bold != null)
                target.Bold = source.Bold;
        }
    }
}
