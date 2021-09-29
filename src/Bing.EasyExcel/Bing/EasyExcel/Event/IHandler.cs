namespace Bing.EasyExcel.Event
{
    /// <summary>
    /// 处理器
    /// </summary>
    public interface IHandler
    {
        /// <summary>
        /// 排序
        /// </summary>
        int Order { get; }
    }
}
