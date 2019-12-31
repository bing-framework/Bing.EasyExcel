using System.IO;
using Bing.EasyExcel.Metadata.Formulas.State;

namespace Bing.EasyExcel.Metadata.Formulas
{
    /// <summary>
    /// 公式上下文
    /// </summary>
    public class FormulaContext
    {
        /// <summary>
        /// 公式文本
        /// </summary>
        public string FormulaText { get; private set; }

        /// <summary>
        /// 读取器
        /// </summary>
        public TextReader Reader { get; private set; }

        /// <summary>
        /// 公式
        /// </summary>
        public Formula Formula { get; private set; }

        /// <summary>
        /// 状态
        /// </summary>
        public IState State { get; set; } = StateHolder.FormulaState;

        /// <summary>
        /// 初始化一个<see cref="FormulaContext"/>类型的实例
        /// </summary>
        /// <param name="formula">公式</param>
        public FormulaContext(string formula)
        {
            FormulaText = formula;
            Reader = new StringReader(formula);
            Formula = new Formula();
        }

        /// <summary>
        /// 解析
        /// </summary>
        public void Parse()
        {
            while (!(State is EndState))
            {
                var i = Reader.Read();
                State.Switch(this, i);
                State.Handle(this, (char)i);
            }
        }
    }
}
