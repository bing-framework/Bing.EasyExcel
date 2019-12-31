using Bing.EasyExcel.Npoi.Metadata;

namespace Bing.EasyExcel.Npoi.Extensions
{
    /// <summary>
    /// NPOI工作簿(<see cref="NPOI.SS.UserModel.IWorkbook"/>) 扩展
    /// </summary>
    public static class WorkbookExtensions
    {
        /// <summary>
        /// 获取适配器
        /// </summary>
        /// <param name="workbook">NPOI工作簿</param>
        public static Workbook GetAdapter(this NPOI.SS.UserModel.IWorkbook workbook) =>
            null == workbook ? null : new Workbook(workbook);
    }
}
