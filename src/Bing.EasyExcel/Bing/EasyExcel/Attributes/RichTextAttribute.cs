using System;

namespace Bing.EasyExcel.Attributes
{
    /// <summary>
    /// 富文本特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    [Serializable]
    public class RichTextAttribute : ExcelAttribute
    {
    }
}
