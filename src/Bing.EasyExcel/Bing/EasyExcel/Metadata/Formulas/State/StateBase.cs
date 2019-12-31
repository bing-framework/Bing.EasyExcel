namespace Bing.EasyExcel.Metadata.Formulas.State
{
    /// <summary>
    /// 状态基类
    /// </summary>
    internal abstract class StateBase : IState
    {
        /// <summary>
        /// 处理
        /// </summary>
        /// <param name="context">公式上下文</param>
        /// <param name="c">字符</param>
        public abstract void Handle(FormulaContext context, char c);

        /// <summary>
        /// 交换
        /// </summary>
        /// <param name="context">公式上下文</param>
        /// <param name="i">索引</param>
        public void Switch(FormulaContext context, int i)
        {
            switch (i)
            {
                case -1:
                    SwitchWhenEndCharacter(context);
                    break;
                case '"':
                    SwitchWhenEscapeCharacter(context);
                    break;
                default:
                    SwitchWhenOtherCharacter(context);
                    break;
            }
        }

        /// <summary>
        /// 交换。当字符为结束符时
        /// </summary>
        /// <param name="context">公式上下文</param>
        public abstract void SwitchWhenEndCharacter(FormulaContext context);

        /// <summary>
        /// 交换。当字符为双引号时
        /// </summary>
        /// <param name="context">公式上下文</param>
        public abstract void SwitchWhenEscapeCharacter(FormulaContext context);

        /// <summary>
        /// 交换。当字符为其它字符时
        /// </summary>
        /// <param name="context">公式上下文</param>
        public abstract void SwitchWhenOtherCharacter(FormulaContext context);
    }
}
