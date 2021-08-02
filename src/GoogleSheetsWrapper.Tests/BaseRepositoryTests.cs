using GoogleSheetsWrapper.Tests.TestObjects;
using NUnit.Framework;
using System;
using System.Collections.Generic;

namespace GoogleSheetsWrapper.Tests
{
    public class BaseRepositoryTests
    {
        [SetUp]
        public void Setup()
        {
        }


        [Test]
        public void VerifySchemaMatches()
        {
            List<object> sampleHeader = new List<object>()
            {
                "Name",
                "Number",
                "Previous Donation Amount",
                "Date"
            };

            var sheetsHelper = new SheetHelper<TestRecord>("", "", "");
            var repo = new TestRepository(sheetsHelper);

            var result = repo.ValidateSchema(sampleHeader);

            Assert.AreEqual(true, result.IsValid);
        }

        [Test]
        public void VerifySchemaDoesNotMatch()
        {
            List<object> sampleHeader = new List<object>()
            {
                "Name",
                "Previous Donation Amount"
            };

            var sheetsHelper = new SheetHelper<TestRecord>("", "", "");
            var repo = new TestRepository(sheetsHelper);

            var result = repo.ValidateSchema(sampleHeader);

            Assert.AreEqual(false, result.IsValid);
        }

        [Test]
        public void Test_Record_Creation()
        {
            List<object> row = new List<object>()
            {
                "Steve",
                "+1(703)-999-2222",
                "$ 100.00",
                "33.625"  // DateTime in serial format for the date time of February 1, 1900 at 3:00 PM
            };

            var record = new TestRecord(row, 1);

            Assert.AreEqual("Steve", record.Name);
            Assert.AreEqual(7039992222, record.PhoneNumber);
            Assert.AreEqual(100, record.DonationAmount);

            var dt = new DateTime(1900, 2, 1, 15, 0, 0);
            Assert.AreEqual(dt, record.DateTime);
        }
    }
}