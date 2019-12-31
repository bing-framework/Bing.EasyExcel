using System.Text;

namespace Bing.EasyExcel.Metadata.Formulas
{
    /// <summary>
    /// 局部
    /// </summary>
    public class Part
    {
        /// <summary>
        /// 字符串拼接器
        /// </summary>
        private readonly StringBuilder _builder;

        /// <summary>
        /// 局部类型
        /// </summary>
        public PartType Type { get; private set; }

        /// <summary>
        /// 初始化一个<see cref="Part"/>类型的实例
        /// </summary>
        /// <param name="type">局部类型</param>
        public Part(PartType type)
        {
            Type = type;
            _builder = new StringBuilder();
        }

        /// <summary>
        /// 追加字符
        /// </summary>
        /// <param name="c">字符</param>
        public void Append(char c) => _builder.Append(c);

        /// <summary>
        /// 输出字符串
        /// </summary>
        public override string ToString() => _builder.ToString();
    }
}
