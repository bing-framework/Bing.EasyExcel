﻿using System;

namespace Bing.EasyExcel
{
    /// <summary>
    /// 列名 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class ColumnNameAttribute : Attribute
    {
        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ColumnNameAttribute"/>类型的实例
        /// </summary>
        /// <param name="name">名称</param>
        public ColumnNameAttribute(string name) => Name = name;
    }
}
