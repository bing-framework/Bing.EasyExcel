using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace Bing.EasyExcel.Metadata.Formulas
{
    /// <summary>
    /// 公式
    /// </summary>
    public class Formula : IEnumerable<Part>
    {
        /// <summary>
        /// 局部列表
        /// </summary>
        private readonly LinkedList<Part> _parts = new LinkedList<Part>();

        /// <summary>
        /// 当前局部
        /// </summary>
        private Part _currentPart;

        /// <summary>
        /// 追加字符
        /// </summary>
        /// <param name="type">局部类型</param>
        /// <param name="c">字符</param>
        public void Append(PartType type, char c)
        {
            if (_currentPart == null || !_currentPart.Type.Equals(type))
            {
                _currentPart = new Part(type);
                _parts.AddLast(_currentPart);
            }
            _currentPart.Append(c);
        }

        /// <summary>
        /// 输出字符串
        /// </summary>
        public override string ToString() => this.ToString(part => part.ToString());

        /// <summary>
        /// 转换为字符串
        /// </summary>
        /// <param name="func">函数</param>
        public string ToString(Func<Part, string> func)
        {
            var builder = new StringBuilder();
            foreach (var part in _parts)
                builder.Append(func(part));
            return builder.ToString();
        }

        /// <summary>
        /// 获取迭代集合
        /// </summary>
        public IEnumerator<Part> GetEnumerator() => this._parts.GetEnumerator();

        /// <summary>
        /// 获取迭代集合
        /// </summary>
        IEnumerator IEnumerable.GetEnumerator() => this._parts.GetEnumerator();
    }
}
