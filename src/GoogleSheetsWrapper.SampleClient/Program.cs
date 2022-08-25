using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest;

namespace GoogleSheetsWrapper.SampleClient
{
    class Program
    {
        static void Main(string[] args)
        {
            var config = BuildConfig();

            // Get the Google Spreadsheet Config Values
            var serviceAccount = config["GOOGLE_SERVICE_ACCOUNT"];
            var documentId = config["GOOGLE_SPREADSHEET_ID"];
            var jsonCredsPath = config["GOOGLE_JSON_CREDS_PATH"];

            // In this case the json creds file is stored locally, but you can store this however you want to (Azure Key Vault, HSM, etc)
            var jsonCredsContent = File.ReadAllText(jsonCredsPath);

            // Create a new SheetHelper class
            var sheetHelper = new SheetHelper(documentId, serviceAccount, "");
            sheetHelper.Init(jsonCredsContent);

            // Get all the rows for the first 2 columns in the spreadsheet
            var rows = sheetHelper.GetRows(new SheetRange("", 1, 1, 2),
                ValueRenderOptionEnum.FORMATTEDVALUE,
                DateTimeRenderOptionEnum.FORMATTEDSTRING);

            // Write all the values from the result set
            foreach (var row in rows)
            {
                foreach (var col in row)
                {
                    Console.Write($"{col}\t");
                }
                Console.Write("\n");
            }

            var appender = new SheetAppender(sheetHelper);

            // Appends weakly typed rows to the spreadsheeet
            appender.AppendRows(new List<List<string>>()
            {
                new List<string>(){"7/1/2022", "abc"},
                new List<string>(){"8/1/2022", "def"}
            });

            // Get all the rows for the first 2 columns in the spreadsheet
            var rows2 = sheetHelper.GetRows(new SheetRange("", 1, 1, 2),
                ValueRenderOptionEnum.FORMATTEDVALUE,
                DateTimeRenderOptionEnum.FORMATTEDSTRING);

            // Write all the values from the result set
            foreach (var row in rows2)
            {
                foreach (var col in row)
                {
                    Console.Write($"{col}\t");
                }
                Console.Write("\n");
            }
        }

        /// <summary>
        /// Really simple method to build the config locally.  This requires you to setup User Secrets locally with Visual Studio
        /// </summary>
        /// <returns></returns>
        private static IConfigurationRoot BuildConfig()
        {
            var devEnvironmentVariable = Environment.GetEnvironmentVariable("NETCORE_ENVIRONMENT");

            var isDevelopment = string.IsNullOrEmpty(devEnvironmentVariable) ||
                                devEnvironmentVariable.ToLower() == "development";

            var builder = new ConfigurationBuilder();
            // tell the builder to look for the appsettings.json file
            builder
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            //only add secrets in development
            if (isDevelopment)
            {
                builder.AddUserSecrets<Program>();
            }

            return builder.Build();
        }
    }
}
