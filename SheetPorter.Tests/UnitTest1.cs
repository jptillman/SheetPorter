using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SheetPorter.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        [DeploymentItem(@".\Fixtures\test.xls", @"Files")]
        public void PorterCorrectlyMapsTableAttributes()
        {
            var model = Porter.Port<Import>(File.OpenRead(@".\Files\test.xls")).ToArray();
            Assert.AreEqual("Jack Jackson", model[0].Name);
            Assert.AreEqual("Tallahassee, FL", model[0].City);
            Assert.AreEqual(DateTime.Parse("04/04/1995"), model[0].GraduationDate);
        }

        [TestMethod]
        [DeploymentItem(@".\Fixtures\test.xls", @"Files")]
        public void PorterCorrectlyMapsPropertySetAttributes()
        {
            var model = Porter.Port<Properties>(File.OpenRead(@".\Files\test.xls")).ToArray();
            Assert.AreEqual("James Jameson", model[0].Name);
            Assert.AreEqual("Seattle, WA", model[0].City);
            Assert.AreEqual(DateTime.Parse("3/5/1980"), model[0].GraduationDate);
        }

        [TestMethod]
        [DeploymentItem(@".\Fixtures\test.xls", @"Files")]
        public void PorterCorrectlyFiltersSubTables()
        {
            var model = Porter.Port<Properties>(File.OpenRead(@".\Files\test.xls")).ToArray();
            Assert.AreEqual("James Jameson", model[0].Name);
            Assert.AreEqual("Seattle, WA", model[0].City);
            Assert.AreEqual(DateTime.Parse("3/5/1980"), model[0].GraduationDate);
            Assert.AreEqual(4, model[0].Grades.Count);
            Assert.AreEqual(3, model[0].AllStudents.Count);
        }

        [TestMethod]
        [DeploymentItem(@".\Fixtures\test.xls", @"Files")]
        public void PorterCorrectlyPopulatesSubproperty()
        {
            var model = Porter.Port<Properties>(File.OpenRead(@".\Files\test.xls")).ToArray();
            Assert.AreEqual("Blue", model[0].ExtraInformation.FavoriteColor);
        }
    }
}
