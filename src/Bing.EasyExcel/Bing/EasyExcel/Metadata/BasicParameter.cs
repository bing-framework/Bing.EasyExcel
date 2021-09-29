using System;
using System.Collections.Generic;

namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 基础参数
    /// </summary>
    public class BasicParameter
    {
        /// <summary>
        /// 标题头
        /// </summary>
        public List<List<string>> Head { get; set; }

        /// <summary>
        /// 类型
        /// </summary>
        public Type Type { get; set; }

        /// <summary>
        /// 自动修剪（工作表名称、内容）
        /// </summary>
        public bool AutoTrim { get; set; }
    }
}
