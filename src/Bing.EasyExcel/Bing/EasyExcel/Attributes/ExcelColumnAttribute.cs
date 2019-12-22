using System;

namespace Bing.EasyExcel.Attributes
{
    /// <summary>
    /// Excel列特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    [Serializable]
    public class ExcelColumnAttribute : ExcelAttribute
    {
        /// <summary>
        /// 列索引
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// 列名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ExcelColumnAttribute"/>类型的实例
        /// </summary>
        public ExcelColumnAttribute() : this(-1) { }

        /// <summary>
        /// 初始化一个<see cref="ExcelColumnAttribute"/>类型的实例
        /// </summary>
        /// <param name="index">列索引</param>
        public ExcelColumnAttribute(int index) => Index = index;

        /// <summary>
        /// 初始化一个<see cref="ExcelColumnAttribute"/>类型的实例
        /// </summary>
        /// <param name="name">列名</param>
        public ExcelColumnAttribute(string name) : this(-1, name) { }

        /// <summary>
        /// 初始化一个<see cref="ExcelColumnAttribute"/>类型的实例
        /// </summary>
        /// <param name="index">列索引</param>
        /// <param name="name">列名</param>
        public ExcelColumnAttribute(int index, string name)
        {
            Index = index;
            Name = name;
        }
    }
}
