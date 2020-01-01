using System.IO;

namespace Bing.EasyExcel.Npoi.Extensions
{
    /// <summary>
    /// NPOI工作簿(<see cref="NPOI.SS.UserModel.IWorkbook"/>) 扩展
    /// </summary>
    public static class WorkbookExtensions
    {
        /// <summary>
        /// 将工作簿转换为二进制文件流
        /// </summary>
        /// <param name="workbook">NPOI工作簿</param>
        public static byte[] SaveToBuffer(this NPOI.SS.UserModel.IWorkbook workbook)
        {
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                return ms.ToArray();
            }
        }
    }
}
