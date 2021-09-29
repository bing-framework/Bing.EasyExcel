using System;
using System.Collections.Generic;
using System.Text;

namespace Bing.EasyExcel.Metadata
{
    public abstract class AbstractParameterBuilder<T, TC> where T : AbstractParameterBuilder<T, TC> where TC : BasicParameter
    {
        public T Head(List<List<string>> head)
        {
            Parameter().Head = head;
            return This();
        }

        public T Head<TType>() => Head(typeof(TType));

        public T Head(Type type)
        {
            Parameter().Type = type;
            return This();
        }

        public T AutoTrim(bool autoTrim)
        {
            Parameter().AutoTrim = autoTrim;
            return This();
        }

        /// <summary>
        /// 返回自身
        /// </summary>
        protected T This() => (T)this;

        /// <summary>
        /// 获取参数
        /// </summary>
        protected abstract TC Parameter();
    }
}
