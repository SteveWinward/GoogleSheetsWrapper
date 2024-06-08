using Microsoft.Extensions.Configuration;

namespace GoogleSheetsWrapper.IntegrationTests
{
    public class EnvironmentVariableConfig
    {
        public string GoogleServiceAccount { get; private set; } = "";

        public string GoogleSpreadsheetId { get; private set; } = "";

        public string JsonCredentials { get; private set; } = "";

        private readonly IConfigurationRoot Config;

        public EnvironmentVariableConfig()
        {
            Config = BuildConfig();

#pragma warning disable CS8601 // Possible null reference assignment.

            GoogleServiceAccount = Config["GOOGLE_SERVICE_ACCOUNT"];
            GoogleSpreadsheetId = Config["GOOGLE_SPREADSHEET_ID"];
            JsonCredentials = Config["GOOGLE_JSON_CREDS"];

#pragma warning restore CS8601 // Possible null reference assignment.
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

            _ = builder.AddEnvironmentVariables();

            //only add secrets in development
            if (isDevelopment)
            {
                _ = builder.AddUserSecrets<EnvironmentVariableConfig>();
            }

            return builder.Build();
        }
    }
}
