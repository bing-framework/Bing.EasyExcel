using System;

namespace Bing.EasyExcel.Attributes
{
    /// <summary>
    /// 公式结果特性。指定应映射公式，这仅适用于字符串属性，对于其它类型，则会直接映射结果
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public class FormulaResultAttribute : ExcelAttribute
    {
    }
}
