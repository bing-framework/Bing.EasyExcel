using System.Drawing;

namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 表格样式
    /// </summary>
    public class TableStyle
    {
        /// <summary>
        /// 表头背景颜色
        /// </summary>
        public Color TableHeadBackgroundColor { get; set; }

        /// <summary>
        /// 表头字体样式
        /// </summary>
        public Font TableHeadFont { get; set; }

        /// <summary>
        /// 表格内容字体样式
        /// </summary>
        public Font TableContentFont { get; set; }

        /// <summary>
        /// 表格内容背景颜色
        /// </summary>
        public Color TableContentBackgroundColor { get; set; }
    }
}
