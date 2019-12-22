using System;

namespace Bing.EasyExcel.Attributes.Write.Styles
{
    /// <summary>
    /// 表格样式特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    [Serializable]
    public class TableStyleAttribute : ExcelAttribute
    {
        /// <summary>
        /// 默认行高。0:不作修改
        /// </summary>
        public int RowDefaultHeight { get; set; }

        /// <summary>
        /// 默认列宽。0:不作修改
        /// </summary>
        public int ColumnDefaultWidth { get; set; }

        /// <summary>
        /// 默认单元格样式
        /// </summary>
        public string CellDefaultStyle { get; set; }

        /// <summary>
        /// 是否显示列头
        /// </summary>
        public bool HeaderVisible { get; set; } = true;

        /// <summary>
        /// 列头高度。如果没有设置，则将使用<see cref="RowDefaultHeight"/>
        /// </summary>
        public int HeaderHeight { get; set; }

        /// <summary>
        /// 列头默认样式。如果没有设置，则将使用<see cref="CellDefaultStyle"/>
        /// </summary>
        public string HeaderDefaultStyle { get; set; }
    }
}
