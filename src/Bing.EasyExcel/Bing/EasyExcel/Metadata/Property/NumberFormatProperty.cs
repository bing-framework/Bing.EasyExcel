using System;

namespace Bing.EasyExcel.Metadata.Property
{
    /// <summary>
    /// 数字格式化属性
    /// </summary>
    public class NumberFormatProperty
    {
        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 舍入方式
        /// </summary>
        public MidpointRounding MidpointRounding { get; set; }
    }
}
