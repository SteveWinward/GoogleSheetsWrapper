using GoogleSheetsWrapper.Utils;
using NUnit.Framework;

namespace CAR.Core.Tests
{
    public class CurrencyParsingTests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void ParseCurrencyWithDollarSign()
        {
            var result = CurrencyParsing.ParseCurrencyString("$100.00");

            Assert.AreEqual(100, result);
        }

        [Test]
        public void ParseCurrencyWithoutDollarSign()
        {
            var result = CurrencyParsing.ParseCurrencyString("100.00");

            Assert.AreEqual(100, result);
        }

        [Test]
        public void ParseCurrencyWithDollarSignAndWhiteSpace()
        {
            var result = CurrencyParsing.ParseCurrencyString("$  100.00");

            Assert.AreEqual(100, result);
        }
    }
}