using System;
using System.Collections.Generic;
using System.Text;
using Bing.EasyExcel.Exports.Metadata;

namespace Bing.EasyExcel.Exports.Builder
{
    public class ExcelWriterBuilder:AbstractExcelWriterParameterBuilder<ExcelWriterBuilder,WriteWorkbook>
    {
        private readonly WriteWorkbook _writeWorkbook;

        public ExcelWriterBuilder() => _writeWorkbook = new WriteWorkbook();

        /// <summary>
        /// 获取参数
        /// </summary>
        protected override WriteWorkbook Parameter() => _writeWorkbook;
    }
}
