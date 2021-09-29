namespace Bing.EasyExcel.Metadata
{
    /// <summary>
    /// 配置持有者
    /// </summary>
    public interface IConfigurationHolder : IHolder
    {
        /// <summary>
        /// 是否新的
        /// </summary>
        bool IsNew { get; }

        /// <summary>
        /// 全局配置
        /// </summary>
        GlobalConfiguration Configuration { get; }
    }
}
