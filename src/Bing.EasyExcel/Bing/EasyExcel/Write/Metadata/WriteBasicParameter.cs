using System.Collections.Generic;
using Bing.EasyExcel.Metadata;

namespace Bing.EasyExcel.Write.Metadata
{
    /// <summary>
    /// 写入基础参数
    /// </summary>
    public class WriteBasicParameter : BasicParameter
    {
        /// <summary>
        /// 相对标题行索引。相对于工作表的现有内容写入头部，索引从0开始。
        /// </summary>
        public int RelativeHeadRowIndex { get; set; }

        /// <summary>
        /// 是否需要标题行
        /// </summary>
        public bool NeedHead { get; set; }

        /// <summary>
        /// 是否使用默认样式。默认：true
        /// </summary>
        public bool UseDefaultStyle { get; set; } = true;

        /// <summary>
        /// 是否自动合并标题行。默认：true
        /// </summary>
        public bool AutomaticMergeHead { get; set; } = true;

        /// <summary>
        /// 忽略自定义列（列索引）
        /// </summary>
        public List<int> ExcludeColumnIndexes { get; set; }

        /// <summary>
        /// 忽略自定义列（字段名）
        /// </summary>
        public List<string> ExcludeColumnFieldNames { get; set; }

        /// <summary>
        /// 只输出自定义列（列索引）
        /// </summary>
        public List<int> IncludeColumnIndexes { get; set; }

        /// <summary>
        /// 只输出自定义列（字段名）
        /// </summary>
        public List<string> IncludeColumnFieldNames { get; set; }
    }
}
