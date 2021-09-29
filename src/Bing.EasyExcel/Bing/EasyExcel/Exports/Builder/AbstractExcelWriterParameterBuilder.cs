using System.Collections.Generic;
using Bing.EasyExcel.Exports.Metadata;
using Bing.EasyExcel.Metadata;
using Bing.EasyExcel.Write.Metadata;

namespace Bing.EasyExcel.Exports.Builder
{
    public abstract class AbstractExcelWriterParameterBuilder<T, TC> : AbstractParameterBuilder<T, TC>
        where T : AbstractExcelWriterParameterBuilder<T, TC>
        where TC : WriteBasicParameter
    {
        public T RelativeHeadRowIndex(int relativeHeadRowIndex)
        {
            Parameter().RelativeHeadRowIndex = relativeHeadRowIndex;
            return This();
        }

        public T NeedHead(bool needHead)
        {
            Parameter().NeedHead = needHead;
            return This();
        }

        public T UseDefaultStyle(bool useDefaultStyle)
        {
            Parameter().UseDefaultStyle = useDefaultStyle;
            return This();
        }

        public T AutomaticMergeHead(bool automaticMergeHead)
        {
            Parameter().AutomaticMergeHead = automaticMergeHead;
            return This();
        }

        public T ExcludeColumnIndexes(List<int> excludeColumnIndexes)
        {
            Parameter().ExcludeColumnIndexes = excludeColumnIndexes;
            return This();
        }

        public T ExcludeColumnFieldNames(List<string> excludeColumnFieldNames)
        {
            Parameter().ExcludeColumnFieldNames = excludeColumnFieldNames;
            return This();
        }

        public T IncludeColumnIndexes(List<int> includeColumnIndexes)
        {
            Parameter().IncludeColumnIndexes = includeColumnIndexes;
            return This();
        }

        public T IncludeColumnFieldNames(List<string> includeColumnFieldNames)
        {
            Parameter().IncludeColumnFieldNames = includeColumnFieldNames;
            return This();
        }
    }
}
