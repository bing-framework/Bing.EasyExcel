using System.Reflection;
using Bing.EasyExcel.Exports.Attributes;

namespace Bing.EasyExcel.Exports
{
    /// <summary>
    /// 导出列属性
    /// </summary>
    public class ExportColumnProperty
    {
        /// <summary>
        /// 最小宽度
        /// </summary>
        public int MinWidth { get; set; }

        /// <summary>
        /// 最大宽度
        /// </summary>
        public int MaxWidth { get; set; }

        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 表头样式
        /// </summary>
        public IBaseStyle HeaderStyle { get; set; } = new CellStyle();

        /// <summary>
        /// 内容样式
        /// </summary>
        public IBaseStyle ColumnStyle { get; set; } = new CellStyle();

        /// <summary>
        /// 字符串格式化
        /// </summary>
        public string StringFormat { get; set; }

        /// <summary>
        /// 单元行合并
        /// </summary>
        public bool RowMerged { get; set; }

        /// <summary>
        /// 是否忽略属性
        /// </summary>
        public bool Ignored { get; set; }

        /// <summary>
        /// 默认值
        /// </summary>
        public object DefaultValue { get; set; }

        /// <summary>
        /// 属性信息
        /// </summary>
        public PropertyInfo PropertyInfo { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ExportColumnProperty"/>类型的实例
        /// </summary>
        /// <param name="property">属性信息</param>
        public ExportColumnProperty(PropertyInfo property)
        {
            Name = property.GetCustomAttribute(typeof(ColumnNameAttribute)) is ColumnNameAttribute columnNameAttr
                ? columnNameAttr.Name
                : property.Name;
            if (property.GetCustomAttribute(typeof(HeaderStyleAttribute)) is HeaderStyleAttribute headerStyleAttr)
                HeaderStyle = headerStyleAttr.Style;
            if (property.GetCustomAttribute(typeof(ColumnStyleAttribute)) is ColumnStyleAttribute columnStyleAttr)
                ColumnStyle = columnStyleAttr.Style;
            if (property.GetCustomAttribute(typeof(StringFormatterAttribute)) is StringFormatterAttribute formatterAttr)
                StringFormat = formatterAttr.Format;
            if (property.GetCustomAttribute(typeof(RowMergedAttribute)) is RowMergedAttribute)
                RowMerged = true;
            if (Helper.HasIgnore(property))
                Ignored = true;
            PropertyInfo = property;
        }
    }
}
