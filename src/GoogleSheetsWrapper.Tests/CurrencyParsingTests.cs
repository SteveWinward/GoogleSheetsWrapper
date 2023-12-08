using GoogleSheetsWrapper.Utils;
using NUnit.Framework;

namespace GoogleSheetsWrapper.Tests
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

            Assert.That(result, Is.EqualTo(100));
        }

        [Test]
        public void ParseCurrencyWithoutDollarSign()
        {
            var result = CurrencyParsing.ParseCurrencyString("100.00");

            Assert.That(result, Is.EqualTo(100));
        }

        [Test]
        public void ParseCurrencyWithDollarSignAndWhiteSpace()
        {
            var result = CurrencyParsing.ParseCurrencyString("$  100.00");

            Assert.That(result, Is.EqualTo(100));
        }
    }
}