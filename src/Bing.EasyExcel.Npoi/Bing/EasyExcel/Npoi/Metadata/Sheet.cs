using System.Collections;
using System.Collections.Generic;
using Bing.EasyExcel.Metadata;
using Bing.EasyExcel.Npoi.Extensions;

namespace Bing.EasyExcel.Npoi.Metadata
{
    /// <summary>
    /// 工作表
    /// </summary>
    public class Sheet : ISheet
    {
        /// <summary>
        /// NPOI工作表
        /// </summary>
        public NPOI.SS.UserModel.ISheet NpoiSheet { get; private set; }

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName => NpoiSheet.SheetName;

        /// <summary>
        /// 获取单元行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        public IRow this[int rowIndex] => NpoiSheet.GetRow(rowIndex).GetAdapter();

        /// <summary>
        /// 初始化一个<see cref="Sheet"/>类型的实例
        /// </summary>
        /// <param name="npoiSheet">NPOI工作表</param>
        public Sheet(NPOI.SS.UserModel.ISheet npoiSheet) => NpoiSheet = npoiSheet;

        /// <summary>
        /// 获取集合
        /// </summary>
        public IEnumerator<IRow> GetEnumerator()
        {
            foreach (NPOI.SS.UserModel.IRow npoiRow in NpoiSheet)
                yield return npoiRow.GetAdapter();
        }

        /// <summary>
        /// 获取集合
        /// </summary>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        /// <summary>
        /// 获取原始值
        /// </summary>
        public object GetOriginal() => NpoiSheet;


        /// <summary>
        /// 复制单元行
        /// </summary>
        /// <param name="start">开始行索引</param>
        /// <param name="end">结束行索引</param>
        public int CopyRows(int start, int end) => NpoiSheet.CopyRows(start, end);

        /// <summary>
        /// 移除单元行
        /// </summary>
        /// <param name="start">开始行索引</param>
        /// <param name="end">结束行索引</param>
        public int RemoveRows(int start, int end) => NpoiSheet.RemoveRows(start, end);
    }
}
