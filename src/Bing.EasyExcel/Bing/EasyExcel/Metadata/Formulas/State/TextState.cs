namespace Bing.EasyExcel.Metadata.Formulas.State
{
    /// <summary>
    /// 文本状态
    /// </summary>
    internal class TextState : StateBase
    {
        /// <summary>
        /// 处理
        /// </summary>
        /// <param name="context">公式上下文</param>
        /// <param name="c">字符</param>
        public override void Handle(FormulaContext context, char c) => context.Formula.Append(PartType.Text, c);

        /// <summary>
        /// 交换。当字符为结束符时
        /// </summary>
        /// <param name="context">公式上下文</param>
        public override void SwitchWhenEndCharacter(FormulaContext context) => throw new FormulaException($"公式异常: {context.FormulaText}");

        /// <summary>
        /// 交换。当字符为双引号时
        /// </summary>
        /// <param name="context">公式上下文</param>
        public override void SwitchWhenEscapeCharacter(FormulaContext context) => context.State = StateHolder.LiteralInferState;

        /// <summary>
        /// 交换。当字符为其它字符时
        /// </summary>
        /// <param name="context">公式上下文</param>
        public override void SwitchWhenOtherCharacter(FormulaContext context)
        {
        }
    }
}
