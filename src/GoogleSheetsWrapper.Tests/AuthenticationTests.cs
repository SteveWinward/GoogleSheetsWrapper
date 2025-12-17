using NUnit.Framework;
using Google.Apis.Auth.OAuth2;

namespace GoogleSheetsWrapper.Tests
{
    [TestFixture]
    public class AuthenticationTests
    {
        [Test]
        public void SheetHelperConstructorWithoutServiceAccountEmailDoesNotThrow()
        {
            // Arrange & Act
            var sheetHelper = new SheetHelper("testSpreadsheetId", "testTab");

            // Assert
            Assert.That(sheetHelper, Is.Not.Null);
            Assert.That(sheetHelper.SpreadsheetID, Is.EqualTo("testSpreadsheetId"));
            Assert.That(sheetHelper.TabName, Is.EqualTo("testTab"));
        }

        [Test]
        public void SheetHelperConstructorWithServiceAccountEmailDoesNotThrow()
        {
            // Arrange & Act
            var sheetHelper = new SheetHelper("testSpreadsheetId", "service@account.com", "testTab");

            // Assert
            Assert.That(sheetHelper, Is.Not.Null);
            Assert.That(sheetHelper.SpreadsheetID, Is.EqualTo("testSpreadsheetId"));
            Assert.That(sheetHelper.ServiceAccountEmail, Is.EqualTo("service@account.com"));
            Assert.That(sheetHelper.TabName, Is.EqualTo("testTab"));
        }

        [Test]
        public void GenericSheetHelperConstructorWithoutServiceAccountEmailDoesNotThrow()
        {
            // Arrange & Act
            var sheetHelper = new SheetHelper<TestObjects.TestRecord>("testSpreadsheetId", "testTab");

            // Assert
            Assert.That(sheetHelper, Is.Not.Null);
            Assert.That(sheetHelper.SpreadsheetID, Is.EqualTo("testSpreadsheetId"));
            Assert.That(sheetHelper.TabName, Is.EqualTo("testTab"));
        }

        [Test]
        public void SheetAppenderConstructorWithoutServiceAccountEmailDoesNotThrow()
        {
            // Arrange & Act
            var appender = new SheetAppender("testSpreadsheetId", "testTab");

            // Assert
            Assert.That(appender, Is.Not.Null);
        }

        [Test]
        public void SheetExporterConstructorWithoutServiceAccountEmailDoesNotThrow()
        {
            // Arrange & Act
            var exporter = new SheetExporter("testSpreadsheetId", "testTab");

            // Assert
            Assert.That(exporter, Is.Not.Null);
        }

        [Test]
        public void SheetHelperInitWithICredentialMethodExists()
        {
            // Arrange
            var sheetHelper = new SheetHelper("testSpreadsheetId", "testTab");

            // Act & Assert - verifies the method signature exists and accepts ICredential
            // Note: We can't fully test Init without a real Google Sheets service
            var initMethod = typeof(SheetHelper).GetMethod("Init", new[] { typeof(ICredential) });
            Assert.That(initMethod, Is.Not.Null, "Init(ICredential) method should exist");
        }

        [Test]
        public void SheetAppenderInitWithICredentialMethodExists()
        {
            // Arrange
            var appender = new SheetAppender("testSpreadsheetId", "testTab");

            // Act & Assert - verifies the method signature exists and accepts ICredential
            var initMethod = typeof(SheetAppender).GetMethod("Init", new[] { typeof(ICredential) });
            Assert.That(initMethod, Is.Not.Null, "Init(ICredential) method should exist");
        }

        [Test]
        public void SheetExporterInitWithICredentialMethodExists()
        {
            // Arrange
            var exporter = new SheetExporter("testSpreadsheetId", "testTab");

            // Act & Assert - verifies the method signature exists and accepts ICredential
            var initMethod = typeof(SheetExporter).GetMethod("Init", new[] { typeof(ICredential) });
            Assert.That(initMethod, Is.Not.Null, "Init(ICredential) method should exist");
        }
    }
}
