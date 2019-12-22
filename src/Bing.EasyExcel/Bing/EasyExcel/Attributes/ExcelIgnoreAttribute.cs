using System;

namespace Bing.EasyExcel.Attributes
{
    /// <summary>
    /// Excel 忽略特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class ExcelIgnoreAttribute : ExcelAttribute
    {
    }
}
