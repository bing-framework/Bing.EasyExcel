using System;
using System.Globalization;
using Bing.EasyExcel.Metadata;
using NPOI.SS.UserModel;

namespace Bing.EasyExcel.Npoi.Extensions
{
    /// <summary>
    /// NPOI单元格(<see cref="NPOI.SS.UserModel.ICell"/>) 扩展
    /// </summary>
    public static partial class CellExtensions
    {
        /// <summary>
        /// 获取单元格的字符串值
        /// </summary>
        /// <param name="cell">NPOI单元格</param>
        public static string GetStringValue(this NPOI.SS.UserModel.ICell cell)
        {
            if (cell == null)
                return string.Empty;
            try
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Error:
                        return cell.ErrorCellValue.ToString();
                    case CellType.Formula:
                        return cell.CellFormula;
                    case CellType.Numeric:
                        return DateUtil.IsCellDateFormatted(cell)
                            ? cell.DateCellValue.ToString(CultureInfo.InvariantCulture)
                            : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                    default:
                        return cell.ToString();
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 设置单元值
        /// </summary>
        /// <param name="cell">NPOI单元格</param>
        /// <param name="value">值</param>
        public static void SetValue(this NPOI.SS.UserModel.ICell cell, object value)
        {
            if (cell == null)
                return;
            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                return;
            }

            var type = value.GetType();
            if (!string.IsNullOrWhiteSpace(type.FullName) && type.FullName.Equals("System.Byte[]"))
            {
                var pictureIndex = cell.Sheet.Workbook.AddPicture(value as byte[], PictureType.PNG);
                var anchor = cell.Sheet.Workbook.GetCreationHelper().CreateClientAnchor();
                anchor.Col1 = cell.ColumnIndex;
                anchor.Col2 = cell.ColumnIndex + cell.GetSpan().ColSpan;
                anchor.Row1 = cell.RowIndex;
                anchor.Row2 = cell.RowIndex + cell.GetSpan().RowSpan;
                var drawing = cell.Sheet.CreateDrawingPatriarch();
                var picture = drawing.CreatePicture(anchor, pictureIndex);
                return;
            }

            switch (Type.GetTypeCode(type))
            {
                case TypeCode.String:
                    cell.SetCellValue(Convert.ToString(value));
                    break;
                case TypeCode.DateTime:
                    cell.SetCellValue(Convert.ToDateTime(value));
                    break;
                case TypeCode.Boolean:
                    cell.SetCellValue(Convert.ToBoolean(value));
                    break;
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Byte:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    cell.SetCellValue(Convert.ToDouble(value));
                    break;
                default:
                    cell.SetCellValue(string.Empty);
                    break;
            }
        }

        /// <summary>
        /// 获取单元格的单元格跨度信息
        /// </summary>
        /// <param name="cell">NPOI单元格</param>
        public static CellSpan GetSpan(this NPOI.SS.UserModel.ICell cell)
        {
            var cellSpan = new CellSpan(1, 1);
            if (cell.IsMergedCell)
            {
                var regionsNum = cell.Sheet.NumMergedRegions;
                for (var i = 0; i < regionsNum; i++)
                {
                    var range = cell.Sheet.GetMergedRegion(i);
                    if (range.FirstRow == cell.RowIndex && range.FirstColumn == cell.ColumnIndex)
                    {
                        cellSpan.RowSpan = range.LastRow - range.FirstRow + 1;
                        cellSpan.ColSpan = range.LastColumn - range.FirstColumn + 1;
                        break;
                    }
                }
            }
            return cellSpan;
        }
    }
}
