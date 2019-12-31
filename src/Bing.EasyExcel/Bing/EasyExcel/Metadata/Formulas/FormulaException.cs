using System;
using Bing.EasyExcel.Exceptions;

namespace Bing.EasyExcel.Metadata.Formulas
{
    /// <summary>
    /// 公式异常
    /// </summary>
    [Serializable]
    public class FormulaException : EasyExcelException
    {
        /// <summary>
        /// 初始化一个<see cref="FormulaException"/>类型的实例
        /// </summary>
        /// <param name="message">错误消息</param>
        public FormulaException(string message) : base(message) { }
    }
}
