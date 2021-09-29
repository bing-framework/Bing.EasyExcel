namespace Bing.EasyExcel.Exports
{
    /// <summary>
    /// Excel导出器
    /// </summary>
    public interface IExcelExporter
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="options">导出选项配置</param>
        byte[] Export(IExportOptions options);
    }
}
