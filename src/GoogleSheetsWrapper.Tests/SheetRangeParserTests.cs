using GoogleSheetsWrapper;
using NUnit.Framework;

namespace CAR.Core.Tests
{
    public class SheetRangeParserTests
    {
        [SetUp]
        public void Setup()
        {

        }

        [Test]
        public void SheetRangeParser_GetTabName_HappyPath()
        {
            var parser = new SheetRangeParser();

            var tabName = parser.GetTabName("MySheet!A1:D4");

            Assert.AreEqual("MySheet", tabName);
        }

        [Test]
        public void SheetRangeParser_IsValidR1C1Notation_HappyPath_With_Tab()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidR1C1Notation("MySheet!R1C1:R2C2");

            Assert.IsTrue(result);
        }

        [Test]
        public void SheetRangeParser_IsValidR1C1Notation_HappyPath_Without_Tab()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidR1C1Notation("R1C1:R2C2");

            Assert.IsTrue(result);
        }

        [Test]
        public void SheetRangeParser_IsValidR1C1Notation_HappyPath_SingleCell_Without_Tab()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidR1C1Notation("R1C1");

            Assert.IsTrue(result);
        }

        [Test]
        public void SheetRangeParser_IsValidR1C1Notation_Fails_With_A1Notation()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidR1C1Notation("A1:D4");

            Assert.IsFalse(result);
        }

        [Test]
        public void SheetRangeParser_IsValidA1Notation_HappyPath_With_Tab()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidA1Notation("MySheet!A1:D4");

            Assert.IsTrue(result);
        }

        [Test]
        public void SheetRangeParser_IsValidA1Notation_HappyPath_With_Tab_Without_Last_Row()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidA1Notation("MySheet!A1:D");

            Assert.IsTrue(result);
        }

        [Test]
        public void SheetRangeParser_IsValidA1Notation_HappyPath_Without_Tab()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidA1Notation("A1:D4");

            Assert.IsTrue(result);
        }
    }
}