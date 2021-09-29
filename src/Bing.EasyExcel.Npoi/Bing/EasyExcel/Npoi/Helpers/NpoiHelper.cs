using System;
using System.IO;
using Bing.EasyExcel.Exports;

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

        /// <summary>
        /// 创建工作簿
        /// </summary>
        /// <param name="format">导出格式</param>
        internal static NPOI.SS.UserModel.IWorkbook CreateWorkbook(ExportFormat format)
        {
            switch (format)
            {
                case ExportFormat.Xlsx:
                    return new NPOI.XSSF.UserModel.XSSFWorkbook();
                case ExportFormat.Xls:
                    return new NPOI.HSSF.UserModel.HSSFWorkbook();
                default:
                    throw new NotImplementedException("excel格式无法解析");
            }
        }
    }
}
