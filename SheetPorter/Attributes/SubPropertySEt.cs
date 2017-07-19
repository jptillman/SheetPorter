using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SheetPorter.Attributes
{
    [AttributeUsage(System.AttributeTargets.Property | System.AttributeTargets.Field)]
    public class SubPropertySet : Attribute
    {
    }
}
