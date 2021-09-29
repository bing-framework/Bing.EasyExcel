namespace Bing.EasyExcel.Metadata.Data
{
    /// <summary>
    /// 数据格式 数据
    /// </summary>
    public class DataFormatData
    {
        /// <summary>
        /// 索引
        /// </summary>
        public short Index { get; set; }

        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 合并
        /// </summary>
        /// <param name="source">源数据</param>
        /// <param name="target">目标数据</param>
        public static void Merge(DataFormatData source, DataFormatData target)
        {
            if (source == null || target == null)
                return;
            if (source.Index >= 0)
                target.Index = source.Index;
            if (!string.IsNullOrWhiteSpace(source.Format))
                target.Format = source.Format;
        }

        /// <summary>
        /// 克隆
        /// </summary>
        public DataFormatData Clone()
        {
            var dataFormatData = new DataFormatData
            {
                Index = Index,
                Format = Format
            };
            return dataFormatData;
        }
    }
}
