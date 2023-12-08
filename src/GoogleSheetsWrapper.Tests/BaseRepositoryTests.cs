using System;
using System.Collections.Generic;
using System.Linq;
using GoogleSheetsWrapper.Tests.TestObjects;
using NUnit.Framework;

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
            var sampleHeader = new List<object>()
            {
                "Name",
                "Number",
                "Price Amount",
                "Date",
                "Quantity"
            };

            var sheetsHelper = new SheetHelper<TestRecord>("", "", "");
            var repo = new TestRepository(sheetsHelper);

            var result = repo.ValidateSchema(sampleHeader);

            Assert.That(result.IsValid, Is.True);
        }

        [Test]
        public void VerifySchemaDoesNotMatch()
        {
            var sampleHeader = new List<object>()
            {
                "Name",
                "Price Amount"
            };

            var sheetsHelper = new SheetHelper<TestRecord>("", "", "");
            var repo = new TestRepository(sheetsHelper);

            var result = repo.ValidateSchema(sampleHeader);

            Assert.That(result.IsValid, Is.False);
        }

        [Test]
        public void TestRecordCreation()
        {
            var row = new List<object>()
            {
                "Steve",
                "+1(703)-999-2222",
                "$ 100.00",
                "33.625",  // DateTime in serial format for the date time of February 1, 1900 at 3:00 PM
                "1,234.56"
            };

            var record = new TestRecord(row, 1);

            Assert.That(record.Name, Is.EqualTo("Steve"));
            Assert.That(record.PhoneNumber, Is.EqualTo(7039992222));
            Assert.That(record.PriceAmount, Is.EqualTo(100));
            Assert.That(record.Quantity, Is.EqualTo(1234.56));

            var dt = new DateTime(1900, 2, 1, 15, 0, 0);
            Assert.That(record.DateTime, Is.EqualTo(dt));
        }

        [Test]
        public void TestRecordCreationColumnOffset()
        {
            var row = new List<object>()
            {
                "Steve",
                "+1(703)-999-2222",
                "$ 100.00",
                "33.625",  // DateTime in serial format for the date time of February 1, 1900 at 3:00 PM
                "1,234.56"
            };

            var attributes = SheetFieldAttributeUtils.GetAllSheetFieldAttributes<TestRecordOffset>();

            var minColumnId = attributes.Min(a => a.Key.ColumnID);

            var record = new TestRecordOffset(row, 1, minColumnId);

            Assert.That(record.Name, Is.EqualTo("Steve"));
            Assert.That(record.PhoneNumber, Is.EqualTo(7039992222));
            Assert.That(record.PriceAmount, Is.EqualTo(100));
            Assert.That(record.Quantity, Is.EqualTo(1234.56));

            var dt = new DateTime(1900, 2, 1, 15, 0, 0);
            Assert.That(record.DateTime, Is.EqualTo(dt));
        }

        [Test]
        public void TestRecordCreationEmptyPhoneNumber()
        {
            var row = new List<object>()
            {
                "Steve",
                "",
                "$ 100.00",
                "33.625",  // DateTime in serial format for the date time of February 1, 1900 at 3:00 PM
                "1,234.56"
            };

            var record = new TestRecord(row, 1);

            Assert.That(record.PhoneNumber, Is.Null);
        }

        [Test]
        public void TestRecordCreationEmptyCurrency()
        {
            var row = new List<object>()
            {
                "Steve",
                "+1(703)-999-2222",
                "",
                "33.625",  // DateTime in serial format for the date time of February 1, 1900 at 3:00 PM
                "1,234.56"
            };

            var record = new TestRecord(row, 1);

            Assert.That(record.PriceAmount, Is.Null);
        }

        [Test]
        public void TestRecordCreationEmptyDateTime()
        {
            var row = new List<object>()
            {
                "Steve",
                "+1(703)-999-2222",
                "$ 100.00",
                "",  // DateTime in serial format for the date time of February 1, 1900 at 3:00 PM
                "1,234.56"
            };

            var record = new TestRecord(row, 1);

            Assert.That(record.DateTime, Is.Null);
        }

        [Test]
        public void TestRecordCreationEmptyDateTimeNonNullable()
        {
            var row = new List<object>()
            {
                "Steve",
                "+1(703)-999-2222",
                "$ 100.00",
                "",  // DateTime in serial format for the date time of February 1, 1900 at 3:00 PM
                "1,234.56"
            };

            var record = new TestRecordNonNullableDateTime(row, 1);

            Assert.That(record.Name, Is.EqualTo("Steve"));
            Assert.That(record.PhoneNumber, Is.EqualTo(7039992222));
            Assert.That(record.PriceAmount, Is.EqualTo(100));
            Assert.That(record.Quantity, Is.EqualTo(1234.56));

            Assert.That(record.DateTime, Is.EqualTo(new DateTime()));
        }
    }
}