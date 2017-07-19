using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SheetPorter.Attributes;

namespace SheetPorter.Tests
{
    [Table("Import", HeaderMatchPattern = "Name")]
    public class Import
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
    }
}
