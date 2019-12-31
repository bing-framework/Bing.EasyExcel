namespace Bing.EasyExcel.Metadata.Formulas.State
{
    /// <summary>
    /// 结束状态
    /// </summary>
    internal class EndState : IState
    {
        /// <summary>
        /// 交换
        /// </summary>
        /// <param name="context">公式上下文</param>
        /// <param name="i">索引</param>
        public void Switch(FormulaContext context, int i)
        {
        }

        /// <summary>
        /// 处理
        /// </summary>
        /// <param name="context">公式上下文</param>
        /// <param name="c">字符</param>
        public void Handle(FormulaContext context, char c)
        {
        }
    }
}
