namespace Bing.EasyExcel.Metadata.Formulas.State
{
    /// <summary>
    /// 文字推断状态
    /// </summary>
    internal class LiteralInferState : StateBase
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
        public override void SwitchWhenEndCharacter(FormulaContext context) => context.State = StateHolder.EndState;

        /// <summary>
        /// 交换。当字符为双引号时
        /// </summary>
        /// <param name="context">公式上下文</param>
        public override void SwitchWhenEscapeCharacter(FormulaContext context) => context.State = StateHolder.TextState;

        /// <summary>
        /// 交换。当字符为其它字符时
        /// </summary>
        /// <param name="context">公式上下文</param>
        public override void SwitchWhenOtherCharacter(FormulaContext context) => context.State = StateHolder.FormulaState;
    }
}
