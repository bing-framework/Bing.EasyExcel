using System;
using System.Collections.Generic;
using System.Text;
using Bing.EasyExcel.Exports;

namespace Bing.EasyExcel.Npoi.Exports
{
    internal class NpoiExcelExporter : IExcelExporter, IDisposable
    {
        /// <summary>
        /// 默认日期格式
        /// </summary>
        private const string DefaultDateFormat = "yyyy-MM-dd HH:mm:ss";

        /// <summary>
        /// 工作簿
        /// </summary>
        private NPOI.SS.UserModel.IWorkbook _workbook;

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="options">导出选项配置</param>
        public byte[] Export(IExportOptions options)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose() => _workbook?.Close();
    }
}
