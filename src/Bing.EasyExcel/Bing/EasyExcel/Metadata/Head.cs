using System.Collections.Generic;
using Bing.EasyExcel.Metadata.Property;

namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 头部
    /// </summary>
    public class Head
    {
        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 字段名
        /// </summary>
        public string FieldName { get; set; }

        /// <summary>
        /// 头部名称列表
        /// </summary>
        public List<string> HeadNameList { get; set; }

        /// <summary>
        /// 是否指定索引
        /// </summary>
        public bool ForceIndex { get; set; }

        /// <summary>
        /// 是否指定名称
        /// </summary>
        public bool ForceName { get; set; }

        /// <summary>
        /// 列宽
        /// </summary>
        public ColumnWidthProperty ColumnWidth { get; set; }

        /// <summary>
        /// 初始化一个<see cref="Head"/>类型的实例
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        /// <param name="fieldName">字段名</param>
        /// <param name="headNameList">头部名称列表</param>
        /// <param name="forceIndex">是否指定索引</param>
        /// <param name="forceName">是否指定名称</param>
        public Head(int columnIndex, string fieldName, List<string> headNameList, bool forceIndex, bool forceName)
        {
            ColumnIndex = columnIndex;
            FieldName = fieldName;
            HeadNameList = headNameList ?? new List<string>();
            ForceName = forceName;
            ForceIndex = forceIndex;
        }
    }
}
