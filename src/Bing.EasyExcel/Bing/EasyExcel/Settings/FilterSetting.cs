namespace Bing.EasyExcel.Settings
{
    /// <summary>
    /// Excel 过滤器设置
    /// </summary>
    internal sealed class FilterSetting
    {
        /// <summary>
        /// 初始化一个<see cref="FilterSetting"/>类型的实例
        /// </summary>
        /// <param name="firstColumn"></param>
        /// <param name="lastColumn"></param>
        public FilterSetting(int firstColumn, int? lastColumn)
        {
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
        }

        /// <summary>
        /// 首列索引
        /// </summary>
        public int FirstColumn { get; }

        /// <summary>
        /// 尾列（最后一列）索引
        /// </summary>
        public int? LastColumn { get; }
    }
}