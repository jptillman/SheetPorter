using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SheetPorter.Attributes;

namespace SheetPorter.Tests
{
    [PropertySet("Extra")]
    public class Extra
    {
        [Column("Favorite color")]
        public string FavoriteColor { get; set; }

        [Column("Hairstyle")]
        public string HairStyle { get; set; }

        [Column("Shoe size")]
        public int ShoeSize { get; set; }
    }
}
