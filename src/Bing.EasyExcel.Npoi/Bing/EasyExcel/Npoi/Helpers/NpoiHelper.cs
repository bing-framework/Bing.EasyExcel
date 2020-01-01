using System.IO;

namespace Bing.EasyExcel.Npoi.Helpers
{
    /// <summary>
    /// NPOI 操作辅助类
    /// </summary>
    public static class NpoiHelper
    {
        /// <summary>
        /// 加载工作簿
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public static NPOI.SS.UserModel.IWorkbook LoadWorkbook(string filePath)
        {
            using (var fileStream=new FileStream(filePath,FileMode.Open,FileAccess.Read))
            {
                return NPOI.SS.UserModel.WorkbookFactory.Create(fileStream);
            }
        }
    }
}
