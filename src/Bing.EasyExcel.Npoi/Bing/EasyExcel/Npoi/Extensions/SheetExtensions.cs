using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Bing.EasyExcel.Metadata.Formulas;
using NPOI.SS.UserModel;

namespace Bing.EasyExcel.Npoi.Extensions
{
    /// <summary>
    /// NPOI工作表(<see cref="NPOI.SS.UserModel.ISheet"/>) 扩展
    /// </summary>
    public static partial class SheetExtensions
    {
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
        /// 移除行
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="rowIndex">行索引</param>
        public static int RemoveRow(this NPOI.SS.UserModel.ISheet sheet, int rowIndex) => sheet.RemoveRows(rowIndex, rowIndex);

        /// <summary>
        /// 移除行
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="startRowIndex">起始行索引</param>
        /// <param name="endRowIndex">结束行索引</param>
        public static int RemoveRows(this NPOI.SS.UserModel.ISheet sheet, int startRowIndex, int endRowIndex)
        {
            var span = endRowIndex - startRowIndex + 1;
            sheet.RemoveMergedRegions(startRowIndex, endRowIndex, null, null);
            sheet.RemovePictures(startRowIndex, endRowIndex, null, null);
            for (var i = endRowIndex; i >= startRowIndex; i--)
            {
                var row = sheet.GetRow(i);
                sheet.RemoveRow(row);
            }
            if (endRowIndex + 1 <= sheet.LastRowNum)
            {
                sheet.ShiftRows(endRowIndex + 1, sheet.LastRowNum, -span, true, false);
                sheet.MovePictures(endRowIndex + 1, null, null, null, moveRowCount: -span);
            }
            return span;
        }

        /// <summary>
        /// 复制行
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="rowIndex">行索引</param>
        public static int CopyRow(this NPOI.SS.UserModel.ISheet sheet, int rowIndex) => sheet.CopyRows(rowIndex, rowIndex);

        /// <summary>
        /// 复制行
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="startRowIndex">起始行索引</param>
        /// <param name="endRowIndex">结束行索引</param>
        public static int CopyRows(this NPOI.SS.UserModel.ISheet sheet, int startRowIndex, int endRowIndex)
        {
            var span = endRowIndex - startRowIndex + 1;
            var newStartRowIndex = startRowIndex + span;
            // 插入空行
            sheet.InsertRows(newStartRowIndex, span);
            // 复制行
            for (var i = startRowIndex; i <= endRowIndex; i++)
            {
                var sourceRow = sheet.GetRow(i);
                var targetRow = sheet.GetRow(i + span);
                targetRow.Height = sourceRow.Height;
                targetRow.ZeroHeight = sourceRow.ZeroHeight;

                #region 复制单元格

                foreach (var sourceCell in sourceRow.Cells)
                {
                    var targetCell = targetRow.GetCell(sourceCell.ColumnIndex);
                    if (null == targetCell)
                        targetCell = targetRow.CreateCell(sourceCell.ColumnIndex);
                    if (null != sourceCell.CellStyle)
                        targetCell.CellStyle = sourceCell.CellStyle;
                    if (null != sourceCell.CellComment)
                        targetCell.CellComment = sourceCell.CellComment;
                    if (null != sourceCell.Hyperlink)
                        targetCell.Hyperlink = sourceCell.Hyperlink;
                    var cfrs = sourceCell.GetConditionalFormattingRules();
                    if (null != cfrs && cfrs.Length > 0)
                        targetCell.AddConditionalFormattingRules(cfrs);
                    targetCell.SetCellType(sourceCell.CellType);
                    // 复制值
                    switch (sourceCell.CellType)
                    {
                        case CellType.Numeric:
                            targetCell.SetCellValue(sourceCell.NumericCellValue);
                            break;
                        case CellType.String:
                            targetCell.SetCellValue(sourceCell.StringCellValue);
                            break;
                        case CellType.Formula:
                            var formula = CopyFormula(sourceCell.CellFormula, span);
                            targetCell.SetCellValue(formula);
                            break;
                        case CellType.Blank:
                            targetCell.SetCellValue(sourceCell.StringCellValue);
                            break;
                        case CellType.Boolean:
                            targetCell.SetCellValue(sourceCell.BooleanCellValue);
                            break;
                        case CellType.Error:
                            targetCell.SetCellValue(sourceCell.ErrorCellValue);
                            break;
                    }
                }

                #endregion
                // 获取模板化内的合并区域
                var regionInfoList = sheet.GetMergedRegionInfos(startRowIndex, endRowIndex, null, null);
                // 复制合并区域
                foreach (var regionInfo in regionInfoList)
                {
                    regionInfo.FirstRow += span;
                    regionInfo.LastRow += span;
                    sheet.AddMergedRegion(regionInfo);
                }
                // 获取模板行内的图片
                var picInfoList = sheet.GetAllPictureInfos(startRowIndex, endRowIndex, null, null);
                // 复制图片
                foreach (var picInfo in picInfoList)
                {
                    picInfo.MaxRow += span;
                    picInfo.MinRow += span;
                    sheet.AddPicture(picInfo);
                }
            }
            return span;
        }

        /// <summary>
        /// 复制公式
        /// </summary>
        /// <param name="formula">公式</param>
        /// <param name="span">跨度</param>
        private static string CopyFormula(string formula, int span)
        {
            var context = new FormulaContext(formula);
            context.Parse();
            return context.Formula.ToString(part =>
            {
                if (part.Type.Equals(PartType.Formula))
                {
                    var regex = new Regex(@"([A-Z]+)(\d+)");
                    return regex.Replace(part.ToString(),
                        m => $"{m.Groups[1].Value}{int.Parse(m.Groups[2].Value) + span}");
                }

                return part.ToString();
            });
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
