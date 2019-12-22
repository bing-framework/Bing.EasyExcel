namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 工作簿加载器
    /// </summary>
    public interface IWorkbookLoader
    {
        /// <summary>
        /// 加载
        /// </summary>
        /// <param name="filePath">文件路径</param>
        IWorkbook Load(string filePath);
    }
}
