using Bing.EasyExcel.Attributes.Write.Styles;

namespace Bing.EasyExcel.Metadata.Property
{
    /// <summary>
    /// 行高属性
    /// </summary>
    public class RowHeightProperty
    {
        /// <summary>
        /// 高度
        /// </summary>
        public short Height { get; set; }

        /// <summary>
        /// 初始化一个<see cref="RowHeightProperty"/>类型的实例
        /// </summary>
        /// <param name="height">高度</param>
        public RowHeightProperty(short height) => Height = height;

        /// <summary>
        /// 构建
        /// </summary>
        /// <param name="headRowHeight">标题行高</param>
        public static RowHeightProperty Build(HeadRowHeightAttribute headRowHeight)
        {
            if (headRowHeight == null || headRowHeight.Height < 0)
                return null;
            return new RowHeightProperty(headRowHeight.Height);
        }

        /// <summary>
        /// 构建
        /// </summary>
        /// <param name="contentRowHeight">内容行高</param>
        public static RowHeightProperty Build(ContentRowHeightAttribute contentRowHeight)
        {
            if (contentRowHeight == null || contentRowHeight.Height < 0)
                return null;
            return new RowHeightProperty(contentRowHeight.Height);
        }
    }
}
