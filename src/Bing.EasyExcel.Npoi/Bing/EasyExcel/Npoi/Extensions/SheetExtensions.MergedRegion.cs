using System.Collections.Generic;
using Bing.EasyExcel.Metadata;
using NPOI.SS.Util;

namespace Bing.EasyExcel.Npoi.Extensions
{
    /// <summary>
    /// NPOI工作表(<see cref="NPOI.SS.UserModel.ISheet"/>) 扩展
    /// </summary>
    public static partial class SheetExtensions
    {
        /// <summary>
        /// 添加合并区域
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="regionInfo">合并区域信息</param>
        public static void AddMergedRegion(this NPOI.SS.UserModel.ISheet sheet, MergedRegionInfo regionInfo)
        {
            var region = new CellRangeAddress(regionInfo.FirstRow, regionInfo.LastRow, regionInfo.FirstCol,
                regionInfo.LastCol);
            sheet.AddMergedRegion(region);
        }

        /// <summary>
        /// 获取工作表中包含合并区域的信息列表
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        public static List<MergedRegionInfo> GetMergedRegionInfos(this NPOI.SS.UserModel.ISheet sheet) => sheet.GetMergedRegionInfos(null, null, null, null);

        /// <summary>
        /// 获取工作表中指定区域包含合并区域的信息列表
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="minRow">最小行索引</param>
        /// <param name="maxRow">最大行索引</param>
        /// <param name="minCol">最小列索引</param>
        /// <param name="maxCol">最大列索引</param>
        public static List<MergedRegionInfo> GetMergedRegionInfos(this NPOI.SS.UserModel.ISheet sheet, int? minRow,
            int? maxRow, int? minCol, int? maxCol)
        {
            var regionInfoList = new List<MergedRegionInfo>();
            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var range = sheet.GetMergedRegion(i);
                if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn, true))
                    regionInfoList.Add(new MergedRegionInfo(i, range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn));
            }
            return regionInfoList;
        }
    }
}
