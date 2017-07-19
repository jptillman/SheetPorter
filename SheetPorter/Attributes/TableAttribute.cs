using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SheetPorter.Attributes
{
    [AttributeUsage(System.AttributeTargets.Class, AllowMultiple = true)]
    public class TableAttribute : Attribute
    {
        public int? Index { get; private set; }
        public string MatchPattern { get; }

        public int? HeaderRowIndex { get; set; }

        public string HeaderMatchPattern { get; set; }

        public TableAttribute(int index)
        {
            Index = index;
        }

        public TableAttribute(string matchPattern)
        {
            MatchPattern = matchPattern;
        }
    }
}
