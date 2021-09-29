using Bing.EasyExcel.Enums;

namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 持有者
    /// </summary>
    public interface IHolder
    {
        /// <summary>
        /// 持有者类型
        /// </summary>
        HolderType Type { get; }
    }
}
