using System;

namespace Bing.EasyExcel.Attributes
{
    /// <summary>
    /// Excel工作表特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class ExcelSheetAttribute : ExcelAttribute
    {
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 工作表索引
        /// </summary>
        public int Index { get; set; } = 0;

        /// <summary>
        /// 是否自动索引
        /// </summary>
        public bool AutoIndex { get; set; } = true;

        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 标题行索引
        /// </summary>
        public int TitleRowIndex { get; set; }

        /// <summary>
        /// 表头行索引
        /// </summary>
        public int HeaderRowIndex { get; set; }

        /// <summary>
        /// 数据行索引
        /// </summary>
        public int DataRowIndex { get; set; }
    }
}
