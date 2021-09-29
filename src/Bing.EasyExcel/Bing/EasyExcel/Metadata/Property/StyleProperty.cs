using Bing.EasyExcel.Attributes.Write.Styles;
using Bing.EasyExcel.Enums.Styles;
using Bing.EasyExcel.Metadata.Data;
using Bing.EasyExcel.Write.Metadata.Styles;

namespace Bing.EasyExcel.Metadata.Property
{
    /// <summary>
    /// 样式属性
    /// </summary>
    public class StyleProperty
    {
        /// <summary>
        /// 数据格式 数据
        /// </summary>
        public DataFormatData DataFormatData { get; set; }

        /// <summary>
        /// 字体样式
        /// </summary>
        public WriteFont WriteFont { get; set; }

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
        public short Rotation { get; set; }

        /// <summary>
        /// 文本缩进的空格数
        /// </summary>
        public short Indent { get; set; }

        /// <summary>
        /// 左边框类型
        /// </summary>
        public BorderStyle BorderLeft { get; set; }

        /// <summary>
        /// 右边框类型
        /// </summary>
        public BorderStyle BorderRight { get; set; }

        /// <summary>
        /// 顶部边框类型
        /// </summary>
        public BorderStyle BorderTop { get; set; }

        /// <summary>
        /// 底部边框类型
        /// </summary>
        public BorderStyle BorderBottom { get; set; }

        /// <summary>
        /// 左边框颜色
        /// </summary>
        public short LeftBorderColor { get; set; }

        /// <summary>
        /// 右边框颜色
        /// </summary>
        public short RightBorderColor { get; set; }

        /// <summary>
        /// 顶部边框颜色
        /// </summary>
        public short TopBorderColor { get; set; }

        /// <summary>
        /// 底部边框颜色
        /// </summary>
        public short BottomBorderColor { get; set; }

        /// <summary>
        /// 填充模式类型
        /// </summary>
        public FillPatternType FillPatternType { get; set; }

        /// <summary>
        /// 背景色
        /// </summary>
        public short FillBackgroundColor { get; set; }

        /// <summary>
        /// 前景色
        /// </summary>
        public short FillForegroundColor { get; set; }

        /// <summary>
        /// 自适应单元格大小
        /// </summary>
        public bool ShrinkToFit { get; set; }

        /// <summary>
        /// 构建
        /// </summary>
        /// <param name="headStyle">标题样式</param>
        public static StyleProperty Build(HeadStyleAttribute headStyle)
        {
            if (headStyle == null)
                return null;
            var styleProperty = new StyleProperty();
            if (headStyle.DataFormat >= 0)
            {
                styleProperty.DataFormatData = new DataFormatData
                {
                    Index = headStyle.DataFormat,
                };
            }

            styleProperty.Hidden = headStyle.Hidden;
            styleProperty.Locked = headStyle.Locked;
            styleProperty.QuotePrefix = headStyle.QuotePrefix;
            styleProperty.HorizontalAlignment = headStyle.HorizontalAlignment;
            styleProperty.Wrapped = headStyle.Wrapped;
            styleProperty.VerticalAlignment = headStyle.VerticalAlignment;
            if (headStyle.Rotation >= 0)
                styleProperty.Rotation = headStyle.Rotation;
            if (headStyle.Indent >= 0)
                styleProperty.Indent = headStyle.Indent;
            styleProperty.BorderLeft = headStyle.BorderLeft;
            styleProperty.BorderRight = headStyle.BorderRight;
            styleProperty.BorderTop = headStyle.BorderTop;
            styleProperty.BorderBottom = headStyle.BorderBottom;
            if (headStyle.LeftBorderColor >= 0)
                styleProperty.LeftBorderColor = headStyle.LeftBorderColor;
            if (headStyle.RightBorderColor >= 0)
                styleProperty.RightBorderColor = headStyle.RightBorderColor;
            if (headStyle.TopBorderColor >= 0)
                styleProperty.TopBorderColor = headStyle.TopBorderColor;
            if (headStyle.BottomBorderColor >= 0)
                styleProperty.BottomBorderColor = headStyle.BottomBorderColor;
            styleProperty.FillPatternType = headStyle.FillPatternType;
            if (headStyle.FillBackgroundColor >= 0)
                styleProperty.FillBackgroundColor = headStyle.FillBackgroundColor;
            if (headStyle.FillForegroundColor >= 0)
                styleProperty.FillForegroundColor = headStyle.FillForegroundColor;
            styleProperty.ShrinkToFit = headStyle.ShrinkToFit;
            return styleProperty;
        }

        /// <summary>
        /// 构建
        /// </summary>
        /// <param name="contentStyle">内容样式</param>
        public static StyleProperty Build(ContentStyleAttribute contentStyle)
        {
            if (contentStyle == null)
                return null;
            var styleProperty = new StyleProperty();
            if (contentStyle.DataFormat >= 0)
            {
                styleProperty.DataFormatData = new DataFormatData
                {
                    Index = contentStyle.DataFormat,
                };
            }

            styleProperty.Hidden = contentStyle.Hidden;
            styleProperty.Locked = contentStyle.Locked;
            styleProperty.QuotePrefix = contentStyle.QuotePrefix;
            styleProperty.HorizontalAlignment = contentStyle.HorizontalAlignment;
            styleProperty.Wrapped = contentStyle.Wrapped;
            styleProperty.VerticalAlignment = contentStyle.VerticalAlignment;
            if (contentStyle.Rotation >= 0)
                styleProperty.Rotation = contentStyle.Rotation;
            if (contentStyle.Indent >= 0)
                styleProperty.Indent = contentStyle.Indent;
            styleProperty.BorderLeft = contentStyle.BorderLeft;
            styleProperty.BorderRight = contentStyle.BorderRight;
            styleProperty.BorderTop = contentStyle.BorderTop;
            styleProperty.BorderBottom = contentStyle.BorderBottom;
            if (contentStyle.LeftBorderColor >= 0)
                styleProperty.LeftBorderColor = contentStyle.LeftBorderColor;
            if (contentStyle.RightBorderColor >= 0)
                styleProperty.RightBorderColor = contentStyle.RightBorderColor;
            if (contentStyle.TopBorderColor >= 0)
                styleProperty.TopBorderColor = contentStyle.TopBorderColor;
            if (contentStyle.BottomBorderColor >= 0)
                styleProperty.BottomBorderColor = contentStyle.BottomBorderColor;
            styleProperty.FillPatternType = contentStyle.FillPatternType;
            if (contentStyle.FillBackgroundColor >= 0)
                styleProperty.FillBackgroundColor = contentStyle.FillBackgroundColor;
            if (contentStyle.FillForegroundColor >= 0)
                styleProperty.FillForegroundColor = contentStyle.FillForegroundColor;
            styleProperty.ShrinkToFit = contentStyle.ShrinkToFit;
            return styleProperty;
        }
    }
}
