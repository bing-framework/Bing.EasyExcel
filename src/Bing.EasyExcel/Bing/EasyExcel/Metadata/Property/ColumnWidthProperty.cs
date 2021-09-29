using Bing.EasyExcel.Attributes.Write.Styles;

namespace Bing.EasyExcel.Metadata.Property
{
    /// <summary>
    /// 列宽属性
    /// </summary>
    public class ColumnWidthProperty
    {
        /// <summary>
        /// 宽度
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ColumnWidthProperty"/>类型的实例
        /// </summary>
        /// <param name="width">宽度</param>
        public ColumnWidthProperty(int width) => Width = width;

        /// <summary>
        /// 构建
        /// </summary>
        /// <param name="columnWidth">列宽</param>
        public static ColumnWidthProperty Build(ColumnWidthAttribute columnWidth)
        {
            if (columnWidth == null || columnWidth.Width < 0)
                return null;
            return new ColumnWidthProperty(columnWidth.Width);
        }
    }
}
