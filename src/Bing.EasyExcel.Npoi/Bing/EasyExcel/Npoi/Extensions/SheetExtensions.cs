using System;
using System.Collections.Generic;
using System.Linq;
using Bing.EasyExcel.Npoi.Metadata;

namespace Bing.EasyExcel.Npoi.Extensions
{
    /// <summary>
    /// NPOI工作表(<see cref="NPOI.SS.UserModel.ISheet"/>) 扩展
    /// </summary>
    public static partial class SheetExtensions
    {
        /// <summary>
        /// 获取适配器
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        internal static Sheet GetAdapter(this NPOI.SS.UserModel.ISheet sheet) => null == sheet ? null : new Sheet(sheet);

        /// <summary>
        /// 插入行
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="rowIndex">行索引</param>
        public static NPOI.SS.UserModel.IRow InsertRow(this NPOI.SS.UserModel.ISheet sheet, int rowIndex) => sheet.InsertRows(rowIndex, 1).FirstOrDefault();

        /// <summary>
        /// 批量插入行
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="rowsCount">插入行数量</param>
        public static NPOI.SS.UserModel.IRow[] InsertRows(this NPOI.SS.UserModel.ISheet sheet, int rowIndex,
            int rowsCount)
        {
            if (rowIndex <= sheet.LastRowNum)
                sheet.ShiftRows(rowIndex, sheet.LastRowNum, rowsCount, true, false);
            var rows = new List<NPOI.SS.UserModel.IRow>();
            for (var i = 0; i < rowsCount; i++)
            {
                var row = sheet.CreateRow(rowIndex + i);
                rows.Add(row);
            }
            return rows.ToArray();
        }

        /// <summary>
        /// 判断指定区域是否在内部或交叉
        /// </summary>
        /// <param name="rangeMinRow">区域最小行索引</param>
        /// <param name="rangeMaxRow">区域最大行索引</param>
        /// <param name="rangeMinCol">区域最小列索引</param>
        /// <param name="rangeMaxCol">区域最大列索引</param>
        /// <param name="targetMinRow">目标最小行索引</param>
        /// <param name="targetMaxRow">目标最大行索引</param>
        /// <param name="targetMinCol">目标最小列索引</param>
        /// <param name="targetMaxCol">目标最大列索引</param>
        /// <param name="onlyInternal">仅在内部</param>
        private static bool IsInternalOrIntersect(int? rangeMinRow, int? rangeMaxRow, int? rangeMinCol,
            int? rangeMaxCol, int targetMinRow, int targetMaxRow, int targetMinCol, int targetMaxCol, bool onlyInternal)
        {
            var tempMinRow = rangeMinRow ?? targetMinRow;
            var tempMaxRow = rangeMaxRow ?? targetMaxRow;
            var tempMinCol = rangeMinCol ?? targetMinCol;
            var tempMaxCol = rangeMaxCol ?? targetMaxCol;
            if (onlyInternal)
            {
                return tempMinRow <= targetMinRow &&
                       tempMaxRow >= targetMaxRow &&
                       tempMinCol <= targetMinCol &&
                       tempMaxCol >= targetMaxCol;
            }

            return Math.Abs(tempMaxRow - tempMinRow) + Math.Abs(targetMaxRow - targetMinRow) >=
                   Math.Abs(tempMaxRow + tempMinRow - targetMaxRow - targetMinRow) &&
                   Math.Abs(tempMaxCol - tempMinCol) + Math.Abs(targetMaxCol - targetMinCol) >=
                   Math.Abs(tempMaxCol + tempMinCol - targetMaxCol - targetMinCol);
        }
    }
}
