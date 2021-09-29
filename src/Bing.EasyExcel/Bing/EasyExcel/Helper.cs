using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;

namespace Bing.EasyExcel
{
    /// <summary>
    /// 帮助类
    /// </summary>
    internal static class Helper
    {
        /// <summary>
        /// 是否有忽略特性
        /// </summary>
        /// <param name="property">属性信息</param>
        public static bool HasIgnore(PropertyInfo property)
        {
            if (property.IsDefined(typeof(NotMappedAttribute)))
                return true;
            if (property.IsDefined(typeof(ExcelIgnoreAttribute)))
                return true;
            return false;
        }
    }
}
