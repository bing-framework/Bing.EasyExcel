namespace Bing.EasyExcel.Metadata.Formulas.State
{
    /// <summary>
    /// 状态持有者
    /// </summary>
    internal static class StateHolder
    {
        /// <summary>
        /// 公式状态
        /// </summary>
        public static IState FormulaState { get; } = new FormulaState();

        /// <summary>
        /// 文本状态
        /// </summary>
        public static IState TextState { get; } = new TextState();

        /// <summary>
        /// 文字推断状态
        /// </summary>
        public static IState LiteralInferState { get; } = new LiteralInferState();

        /// <summary>
        /// 结束状态
        /// </summary>
        public static IState EndState { get; } = new EndState();
    }
}
