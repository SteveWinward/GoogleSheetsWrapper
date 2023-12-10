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

        [TestCase("$100.00", 100)]
        [TestCase("100.00", 100)]
        [TestCase("$  100.00", 100)]
        public void CurrencyTests(string input, double expected)
        {
            var result = CurrencyParsing.ParseCurrencyString(input);

            Assert.That(result, Is.EqualTo(expected));
        }
    }
}