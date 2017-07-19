using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SheetPorter.Attributes
{
    [AttributeUsage(System.AttributeTargets.Class, AllowMultiple = true)]
    public class PropertySetAttribute : Attribute
    {
        public int? Index { get; }
        public string MatchPattern { get; }

        public PropertySetAttribute(int index)
        {
            Index = index;
        }

        public PropertySetAttribute(string matchPattern)
        {
            MatchPattern = matchPattern;

        }
    }
}
