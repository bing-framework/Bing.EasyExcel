using System;
using Bing.EasyExcel.Enums.Styles;

namespace Bing.EasyExcel.Attributes.Write.Styles
{
    /// <summary>
    /// 内容样式 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class ContentStyleAttribute : ExcelAttribute
    {
        /// <summary>
        /// 数据格式
        /// </summary>
        public short DataFormat { get; set; } = -1;

        /// <summary>
        /// 隐藏单元格
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// 锁定单元格
        /// </summary>
        public bool Locked { get; set; }

        /// <summary>
        /// 引用前缀
        /// </summary>
        public bool QuotePrefix { get; set; }

        /// <summary>
        /// 水平对齐方式
        /// </summary>
        public HorizontalAlignment HorizontalAlignment { get; set; }

        /// <summary>
        /// 自动换行
        /// </summary>
        public bool Wrapped { get; set; }

        /// <summary>
        /// 垂直对齐方式
        /// </summary>
        public VerticalAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// 文本的旋转度数
        /// </summary>
        /// <remarks>
        /// 注意：<br />
        /// HSSF=[-90,90]<br />
        /// XSSF=[0,180]
        /// </remarks>
        public short Rotation { get; set; } = -1;

        /// <summary>
        /// 文本缩进的空格数
        /// </summary>
        public short Indent { get; set; } = -1;

        /// <summary>
        /// 左边框类型
        /// </summary>
        public BorderStyle BorderLeft { get; set; } = BorderStyle.None;

        /// <summary>
        /// 右边框类型
        /// </summary>
        public BorderStyle BorderRight { get; set; } = BorderStyle.None;

        /// <summary>
        /// 顶部边框类型
        /// </summary>
        public BorderStyle BorderTop { get; set; } = BorderStyle.None;

        /// <summary>
        /// 底部边框类型
        /// </summary>
        public BorderStyle BorderBottom { get; set; } = BorderStyle.None;

        /// <summary>
        /// 左边框颜色
        /// </summary>
        public short LeftBorderColor { get; set; } = -1;

        /// <summary>
        /// 右边框颜色
        /// </summary>
        public short RightBorderColor { get; set; } = -1;

        /// <summary>
        /// 顶部边框颜色
        /// </summary>
        public short TopBorderColor { get; set; } = -1;

        /// <summary>
        /// 底部边框颜色
        /// </summary>
        public short BottomBorderColor { get; set; } = -1;

        /// <summary>
        /// 填充模式类型
        /// </summary>
        public FillPatternType FillPatternType { get; set; }

        /// <summary>
        /// 背景色
        /// </summary>
        public short FillBackgroundColor { get; set; } = -1;

        /// <summary>
        /// 前景色
        /// </summary>
        public short FillForegroundColor { get; set; } = -1;

        /// <summary>
        /// 自适应单元格大小
        /// </summary>
        public bool ShrinkToFit { get; set; }
    }
}
