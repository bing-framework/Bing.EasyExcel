namespace Bing.EasyExcel.Metadata.Formulas.State
{
    /// <summary>
    /// 状态
    /// </summary>
    public interface IState
    {
        /// <summary>
        /// 交换
        /// </summary>
        /// <param name="context">公式上下文</param>
        /// <param name="i">索引</param>
        void Switch(FormulaContext context, int i);

        /// <summary>
        /// 处理
        /// </summary>
        /// <param name="context">公式上下文</param>
        /// <param name="c">字符</param>
        void Handle(FormulaContext context, char c);
    }
}
