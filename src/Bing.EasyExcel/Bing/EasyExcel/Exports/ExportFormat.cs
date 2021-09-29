using System.ComponentModel;

namespace Bing.EasyExcel.Exports
{
    /// <summary>
    /// 导出格式
    /// </summary>
    public enum ExportFormat
    {
        /// <summary>
        /// 2007
        /// </summary>
        [Description("Excel2007+")]
        Xlsx = 1,
        /// <summary>
        /// 97-2003
        /// </summary>
        [Description("Excel2003")]
        Xls = 2,
    }
}
