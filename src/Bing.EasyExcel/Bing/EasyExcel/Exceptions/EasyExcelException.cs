using System;
using System.Runtime.Serialization;

namespace Bing.EasyExcel.Exceptions
{
    /// <summary>
    /// EasyExcel异常
    /// </summary>
    [Serializable]
    public class EasyExcelException : Exception
    {
        /// <summary>
        /// 初始化一个<see cref="EasyExcelException"/>类型的实例
        /// </summary>
        public EasyExcelException() : base("Office异常") { }

        /// <summary>
        /// 初始化一个<see cref="EasyExcelException"/>类型的实例
        /// </summary>
        /// <param name="msgFormat">格式化消息</param>
        /// <param name="objects">格式化参数</param>
        public EasyExcelException(string msgFormat, params object[] objects) : base(string.Format(msgFormat, objects)) { }

        /// <summary>
        /// 初始化一个<see cref="EasyExcelException"/>类型的实例
        /// </summary>
        /// <param name="message">序列化信息</param>
        /// <param name="innerException">错误来源</param>
        public EasyExcelException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 初始化一个<see cref="EasyExcelException"/>类型的实例
        /// </summary>
        /// <param name="info">序列化信息</param>
        /// <param name="context">流上下文</param>
        public EasyExcelException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
