using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SheetPorter.Attributes
{
    [AttributeUsage(System.AttributeTargets.Property | System.AttributeTargets.Field)]
    public class ColumnAttribute : Attribute
    {
        public int? Index { get; }
        public string MatchPattern { get; }

        public ColumnAttribute(int index)
        {
            Index = index;
        }

        public ColumnAttribute(string matchPattern)
        {
            MatchPattern = matchPattern;
        }
    }
}
