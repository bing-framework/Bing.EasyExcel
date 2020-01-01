using Bing.EasyExcel.Metadata;
using Bing.EasyExcel.Npoi.Extensions;
using Bing.EasyExcel.Npoi.Helpers;

namespace Bing.EasyExcel.Npoi.Metadata
{
    /// <summary>
    /// 工作簿加载器
    /// </summary>
    public class WorkbookLoader : IWorkbookLoader
    {
        /// <summary>
        /// 加载
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public IWorkbook Load(string filePath) => NpoiHelper.LoadWorkbook(filePath).GetAdapter();
    }
}
