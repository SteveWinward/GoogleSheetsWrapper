using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Http;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using NUnit.Framework;

namespace GoogleSheetsWrapper.Tests
{
    public class SheetAppenderTests : IDisposable
    {
        private RecordingHttpMessageHandler _httpHandler;
        private SheetAppender _appender;

        [SetUp]
        public void SetUp()
        {
            _httpHandler = new RecordingHttpMessageHandler();
            var service = new SheetsService(new BaseClientService.Initializer
            {
                ApplicationName = "GoogleSheetsWrapper.Tests",
                HttpClientFactory = new HttpClientFromMessageHandlerFactory(_ =>
                    new HttpClientFromMessageHandlerFactory.ConfiguredHttpMessageHandler(_httpHandler, false, false))
            });

            var sheetHelper = new SheetHelper("spreadsheet-id", "service-account@example.com", "Sheet1")
            {
                SheetID = 42
            };

            typeof(SheetHelper)
                .GetProperty(nameof(SheetHelper.Service), BindingFlags.Instance | BindingFlags.Public)
                .SetValue(sheetHelper, service);

            _appender = new SheetAppender(sheetHelper);
        }

        [Test]
        public void AppendObjectAppendsHeaderAndRecordsInBatches()
        {
            var records = new[]
            {
                new TestRecord { Name = "Alice", Age = 30 },
                new TestRecord { Name = "Bob", Age = 31 },
                new TestRecord { Name = "Charlie", Age = 32 }
            };

            _appender.AppendObject(records, new BatchInformation
            {
                BatchSize = 3,
                BatchWaitTime = 0
            });

            Assert.That(_httpHandler.RequestBodies, Has.Count.EqualTo(2));
            Assert.That(_httpHandler.RequestBodies[0], Does.Contain("Name"));
            Assert.That(_httpHandler.RequestBodies[0], Does.Contain("Alice"));
            Assert.That(_httpHandler.RequestBodies[0], Does.Contain("Bob"));
            Assert.That(_httpHandler.RequestBodies[1], Does.Contain("Charlie"));
        }

        [Test]
        public void AppendCsvAppendsHeaderAndValues()
        {
            using var stream = new System.IO.MemoryStream(
                System.Text.Encoding.UTF8.GetBytes("Name,Age\nAlice,30\n"));

            _appender.AppendCsv(stream, true, new BatchInformation
            {
                BatchSize = 10,
                BatchWaitTime = 0
            });

            Assert.That(_httpHandler.RequestBodies, Has.Count.EqualTo(1));
            Assert.That(_httpHandler.RequestBodies[0], Does.Contain("Name"));
            Assert.That(_httpHandler.RequestBodies[0], Does.Contain("Alice"));
            Assert.That(_httpHandler.RequestBodies[0], Does.Contain("30"));
        }

        [Test]
        public void AppendCsvReadsRecordsIncrementally()
        {
            var csv = "Name,Age\n" + string.Join("\n", Enumerable.Range(0, 10000).Select(index => $"Person{index},{index}"));
            using var stream = new TrackingMemoryStream(System.Text.Encoding.UTF8.GetBytes(csv));

            _appender.AppendCsv(stream, true, new BatchInformation
            {
                BatchSize = 1,
                BatchWaitTime = 0,
                SkipWritingHeaderRow = true,
                ShouldCancel = () =>
                {
                    Assert.That(stream.BytesRead, Is.LessThan(stream.Length));
                    return true;
                }
            });

            Assert.That(_httpHandler.RequestBodies, Has.Count.EqualTo(1));
            Assert.That(_httpHandler.RequestBodies[0], Does.Contain("Person0"));
            Assert.That(_httpHandler.RequestBodies[0], Does.Not.Contain("Person1"));
        }

        [Test]
        public void AppendObjectWhenCancellationCallbackReturnsTrueStopsAfterCurrentBatch()
        {
            var callbackCount = 0;
            var records = new[]
            {
                new TestRecord { Name = "Alice", Age = 30 },
                new TestRecord { Name = "Bob", Age = 31 },
                new TestRecord { Name = "Charlie", Age = 32 }
            };

            _appender.AppendObject(records, new BatchInformation
            {
                BatchSize = 1,
                BatchWaitTime = 0,
                SkipWritingHeaderRow = true,
                ShouldCancel = () =>
                {
                    callbackCount++;
                    return true;
                }
            });

            Assert.That(callbackCount, Is.EqualTo(1));
            Assert.That(_httpHandler.RequestBodies, Has.Count.EqualTo(1));
            Assert.That(_httpHandler.RequestBodies[0], Does.Contain("Alice"));
            Assert.That(_httpHandler.RequestBodies[0], Does.Not.Contain("Bob"));
        }

        [Test]
        public void AppendObjectDoesNotCallCancellationCallbackAfterFinalExactBatch()
        {
            var callbackCount = 0;
            var records = new[]
            {
                new TestRecord { Name = "Alice", Age = 30 },
                new TestRecord { Name = "Bob", Age = 31 },
                new TestRecord { Name = "Charlie", Age = 32 },
                new TestRecord { Name = "Diana", Age = 33 }
            };

            _appender.AppendObject(records, new BatchInformation
            {
                BatchSize = 2,
                BatchWaitTime = 0,
                SkipWritingHeaderRow = true,
                ShouldCancel = () =>
                {
                    callbackCount++;
                    return false;
                }
            });

            Assert.That(_httpHandler.RequestBodies, Has.Count.EqualTo(2));
            Assert.That(callbackCount, Is.EqualTo(1));
        }

        private sealed class TestRecord
        {
            public string Name { get; set; }

            public int Age { get; set; }
        }

        private sealed class RecordingHttpMessageHandler : HttpMessageHandler
        {
            public List<string> RequestBodies { get; } = new List<string>();

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                RequestBodies.Add(ReadRequestBody(request));

                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("{}")
                });
            }

            private static string ReadRequestBody(HttpRequestMessage request)
            {
                if (request.Content == null)
                {
                    return string.Empty;
                }

                using var contentStream = request.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
                using var bodyStream = request.Content.Headers.ContentEncoding.Contains("gzip")
                    ? new GZipStream(contentStream, CompressionMode.Decompress)
                    : contentStream;
                using var reader = new StreamReader(bodyStream);

                return reader.ReadToEnd();
            }
        }

        private sealed class TrackingMemoryStream : MemoryStream
        {
            public TrackingMemoryStream(byte[] buffer)
                : base(buffer)
            {
            }

            public long BytesRead { get; private set; }

            public override int Read(byte[] buffer, int offset, int count)
            {
                var bytesRead = base.Read(buffer, offset, count);
                BytesRead += bytesRead;
                return bytesRead;
            }

            public override int Read(Span<byte> buffer)
            {
                var bytesRead = base.Read(buffer);
                BytesRead += bytesRead;
                return bytesRead;
            }
        }

        public void Dispose()
        {
            _httpHandler.Dispose();
            GC.SuppressFinalize(this);
        }
    }
}
