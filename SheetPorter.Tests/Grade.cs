using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SheetPorter.Attributes;

namespace SheetPorter.Tests
{
    [Table(3, HeaderMatchPattern = "Name")]
    public class Grade 
    {
        [Column("Name")]
        public string Name { get; set; }

        [Column("Semester")]
        public int Semester { get; set; }

        [Column("Grade")]
        public char GradeLetter { get; set; }

    }
}
