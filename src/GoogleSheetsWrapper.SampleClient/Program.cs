using System;
using System.IO;
using Microsoft.Extensions.Configuration;
using GoogleSheetsWrapper.SampleClient.SampleModel;

namespace GoogleSheetsWrapper.SampleClient
{
    public class Program
    {
        public static void Main()
        {
            var config = BuildConfig();

            // Get the Google Spreadsheet Config Values
            var serviceAccount = config["GOOGLE_SERVICE_ACCOUNT"];
            var documentId = config["GOOGLE_SPREADSHEET_ID"];
            var jsonCredsPath = config["GOOGLE_JSON_CREDS_PATH"];

            // In this case the json creds file is stored locally, but you can store this however you want to (Azure Key Vault, HSM, etc)
            var jsonCredsContent = File.ReadAllText(jsonCredsPath);

            // Create a new SheetHelper class
            var sheetHelper = new SheetHelper<SampleRecord>(documentId, serviceAccount, "");
            sheetHelper.Init(jsonCredsContent);

            var repoConfig = new BaseRepositoryConfiguration()
            {
                // Does the table have a header row?
                HasHeaderRow = true,
                // Are there any blank rows before the header row starts?
                HeaderRowOffset = 0,
                // Are there any blank rows before the first row in the data table starts?                
                DataTableRowOffset = 0,
            };

            var respository = new SampleRepository(sheetHelper, repoConfig);

            var records = respository.GetAllRecords();

            foreach (var record in records)
            {
                try
                {
                    Foo(record.TaskName);
                    record.Result = true;
                    record.DateExecuted = DateTime.UtcNow;

                    var result = respository.SaveFields(
                        record,
                        r => r.Result,
                        r => r.DateExecuted);
                }
                catch (Exception ex)
                {
                    record.Result = false;
                    record.ErrorMessage = ex.Message;
                    record.DateExecuted = DateTime.UtcNow;

                    var result = respository.SaveFields(
                        record,
                        r => r.Result,
                        r => r.ErrorMessage,
                        r => r.DateExecuted);
                }
            }
        }

        private static void Foo(string name)
        {
            // Do some operation based on the record
            Console.WriteLine(name);
        }

        /// <summary>
        /// Really simple method to build the config locally.  This requires you to setup User Secrets locally with Visual Studio
        /// </summary>
        /// <returns></returns>
        private static IConfigurationRoot BuildConfig()
        {
            var devEnvironmentVariable = Environment.GetEnvironmentVariable("NETCORE_ENVIRONMENT");

            var isDevelopment = string.IsNullOrEmpty(devEnvironmentVariable) ||
                                devEnvironmentVariable.Equals("development", StringComparison.OrdinalIgnoreCase);

            var builder = new ConfigurationBuilder();
            // tell the builder to look for the appsettings.json file
            _ = builder
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            //only add secrets in development
            if (isDevelopment)
            {
                _ = builder.AddUserSecrets<Program>();
            }

            return builder.Build();
        }
    }
}