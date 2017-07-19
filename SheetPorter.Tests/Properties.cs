using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SheetPorter.Attributes;

namespace SheetPorter.Tests
{
    [PropertySet("Properties")]
    public class Properties
    {
        [Column("Name")]
        public string Name { get; set; }

        [Column("City")]
        public string City { get; set; }

        [Column("Age")]
        public int Age { get; set; }

        [Column("Graddate")]
        public DateTime GraduationDate { get; set; }

        public List<Grade> Grades { get; set; } = new List<Grade>();

        [SubTable(AutoFilter = false)]
        public List<Import> AllStudents { get; set; }

        [SubPropertySet]
        public Extra ExtraInformation { get; set; }
    }
}
