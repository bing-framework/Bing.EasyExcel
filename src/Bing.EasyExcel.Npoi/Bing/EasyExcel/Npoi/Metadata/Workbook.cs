using System.Collections;
using System.Collections.Generic;
using Bing.EasyExcel.Metadata;
using Bing.EasyExcel.Npoi.Extensions;

namespace Bing.EasyExcel.Npoi.Metadata
{
    /// <summary>
    /// 工作簿
    /// </summary>
    public class Workbook : IWorkbook
    {
        /// <summary>
        /// NPOI工作簿
        /// </summary>
        public NPOI.SS.UserModel.IWorkbook NpoiWorkbook { get; private set; }

        /// <summary>
        /// 初始化一个<see cref="Workbook"/>类型的实例
        /// </summary>
        /// <param name="npoiWorkbook">NPOI工作簿</param>
        public Workbook(NPOI.SS.UserModel.IWorkbook npoiWorkbook) => NpoiWorkbook = npoiWorkbook;

        /// <summary>
        /// 获取集合
        /// </summary>
        public IEnumerator<ISheet> GetEnumerator()
        {
            foreach (NPOI.SS.UserModel.ISheet npoiSheet in NpoiWorkbook)
                yield return npoiSheet.GetAdapter();
        }

        /// <summary>
        /// 获取集合
        /// </summary>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        /// <summary>
        /// 获取原始值
        /// </summary>
        public object GetOriginal() => NpoiWorkbook;

        /// <summary>
        /// 获取工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        public ISheet this[string sheetName] => NpoiWorkbook.GetSheet(sheetName).GetAdapter();

        /// <summary>
        /// 保存到流
        /// </summary>
        public byte[] SaveToBuffer() => NpoiWorkbook.SaveToBuffer();
    }
}
