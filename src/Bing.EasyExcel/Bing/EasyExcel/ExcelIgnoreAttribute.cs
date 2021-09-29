using System;

namespace Bing.EasyExcel
{
    /// <summary>
    /// Excel忽略 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelIgnoreAttribute : Attribute
    {
    }
}
