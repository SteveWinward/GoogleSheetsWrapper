# GoogleSheetsWrapper

## Badge Statuses

| Badge Name | Badge Status |
| ---------- | :------------: |
| Current Nuget Version | [![Nuget](https://img.shields.io/nuget/v/GoogleSheetsWrapper)](https://www.nuget.org/packages/GoogleSheetsWrapper/) |
| Total Nuget Downloads | [![Nuget Downloads](https://img.shields.io/nuget/dt/GoogleSheetsWrapper)](https://www.nuget.org/packages/GoogleSheetsWrapper/) |
| Open Source License Details | [![GitHub](https://img.shields.io/github/license/SteveWinward/GoogleSheetsWrapper)](LICENSE) |
| Latest Build Status | ![build status](https://github.com/SteveWinward/GoogleSheetsWrapper/actions/workflows/main_build.yml/badge.svg) |

## Google Sheets API .NET Wrapper Library
> [!IMPORTANT]
> Using this library requires you to use a Service Account to access your Google Sheets spreadsheets.  Please review the authentication section farther down for more details on how to set this up.

This library allows you to use strongly typed objects against a Google Sheets spreadsheet without having to have knowledge on the Google Sheets API methods and protocols. 

The following Google Sheets API operations are supported: 

* Reading all rows
* Appending new rows
* Deleting rows
* Updating specific cells

All operations above are encapsulated in the ````SheetHelper```` class. 

There are also base classes, ````BaseRecord```` and ````BaseRepository```` to simplify transforming raw Google Sheets rows into .NET objects. 

A really simple console application using this library is included in this project below,

[GoogleSheetsWrapper.SampleClient Project](src/GoogleSheetsWrapper.SampleClient/Program.cs)

To setup the sample application, you also need to configure [User Secrets in Visual Studio](https://docs.microsoft.com/en-us/aspnet/core/security/app-secrets?view=aspnetcore-5.0&tabs=windows) to run it locally.

## Extend BaseRecord and BaseRepository

Extending the ````BaseRecord```` class you can decorate properties with the ````SheetFieldAttribute```` to describe the column header name, the column index and the field type (ie ````string````, ````DateTime````, etc)

> [!NOTE]
> The column index is 1 based and not 0 based. The first column 'A' is equivalent to the column ID of 1. 

```csharp
public class TestRecord : BaseRecord
{
    [SheetField(
        DisplayName = "Name",
        ColumnID = 1,
        FieldType = SheetFieldType.String)]
    public string Name { get; set; }

    [SheetField(
        DisplayName = "Number",
        ColumnID = 2,
        FieldType = SheetFieldType.PhoneNumber)]
    public long? PhoneNumber { get; set; }

    [SheetField(
        DisplayName = "Price Amount",
        ColumnID = 3,
        FieldType = SheetFieldType.Currency)]
    public double? PriceAmount { get; set; }

    [SheetField(
        DisplayName = "Date",
        ColumnID = 4,
        FieldType = SheetFieldType.DateTime)]
    public DateTime? DateTime { get; set; }

    [SheetField(
        DisplayName = "Quantity",
        ColumnID = 5,
        FieldType = SheetFieldType.Number)]
    public double? Quantity { get; set; }

    public TestRecord() { }

    // This constructor signature is required to define!
    public TestRecord(IList<object> row, int rowId, int minColumnId = 1)
        : base(row, rowId, minColumnId)
    {
    }
}
```

Extending the ````BaseRepository```` allows you to define your own access layer to the Google Sheets tab you want to work with. 

```csharp
public class TestRepository : BaseRepository<TestRecord>
{
    public TestRepository() { }

    public TestRepository(SheetHelper<TestRecord> sheetsHelper, BaseRepositoryConfiguration config)
        : base(sheetsHelper, config) { }
}
```

## Core Operations (Strongly Typed)

Before you run the following code you will need to setup a Google Service Account and create a service account key.  You also need to decide how to store your environment variables and secrets (ie the Google service account key)

More details can be found [here](#authentication)

```csharp
// You need to implement your own configuration management solution here!
var settings = AppSettings.FromEnvironment();

// Create a SheetHelper class for the specified Google Sheet and Tab name
var sheetHelper = new SheetHelper<TestRecord>(
    // https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID_IS_HERE>/edit#gid=0
    settings.GoogleSpreadsheetId,
    // The email for the service account you created
    settings.GoogleServiceAccountName,
    // the name of the tab you want to access, leave blank if you want the default first tab
    settings.GoogleMainSheetName);

sheetHelper.Init(settings.JsonCredential);

// Create a Repository for the TestRecord class
var repoConfig = new BaseRepositoryConfiguration()
{
    // Does the table have a header row?
    HasHeaderRow = true,
    // Are there any blank rows before the header row starts?
    HeaderRowOffset = 0,
    // Are there any blank rows before the first row in the data table starts?                
    DataTableRowOffset = 0,
};

var repository = new TestRepository(sheetHelper, repoConfig);

// Validate that the header names match the expected format defined with the SheetFieldAttribute values
var result = repository.ValidateSchema();

if(!result.IsValid)
{
    throw new ArgumentException(result.ErrorMessage);
}

// Get all rows from the Google Sheet
var allRecords = repository.GetAllRecords();

// Get the first record
var firstRecord = allRecords.First();

// Update the PriceAmount field and save it back to Google Sheets
firstRecord.PriceAmount = 99.99;
repository.SaveField(firstRecord, (r) => r.PriceAmount);

// Delete the first record from Google Sheets
repository.DeleteRecord(firstRecord);

// Add a new record to the end of the Google Sheets tab
repository.AddRecord(new TestRecord()
{
    Name = "John",
    Number = 2021112222,
    PriceAmount = 250.50,
    Date = DateTime.Now(),
    Quantity = 123
});

```
## Weakly Typed Operations
If you don't want to extend the ````BaseRecord```` class, you can use non typed operations with the ````SheetHelper```` class.  Below is an example of that,

````csharp
// Get the Google Spreadsheet Config Values
var serviceAccount = config["GOOGLE_SERVICE_ACCOUNT"];
var documentId = config["GOOGLE_SPREADSHEET_ID"];
var jsonCredsPath = config["GOOGLE_JSON_CREDS_PATH"];

// In this case the json creds file is stored locally, 
// but you can store this however you want to (Azure Key Vault, HSM, etc)
var jsonCredsContent = File.ReadAllText(jsonCredsPath);

// Create a new SheetHelper class
var sheetHelper = new SheetHelper(documentId, serviceAccount, "");
sheetHelper.Init(jsonCredsContent);

// Append new rows to the spreadsheet
var appender = new SheetAppender(sheetHelper);

// Appends weakly typed rows to the spreadsheeet
appender.AppendRows(new List<List<string>>()
{
    new List<string>(){"7/1/2022", "abc"},
    new List<string>(){"8/1/2022", "def"}
});

// Get all the rows for the first 2 columns in the spreadsheet
var rows = sheetHelper.GetRows(new SheetRange("", 1, 1, 2));

// Print out all the values from the result set
foreach (var row in rows)
{
    foreach (var col in row)
    {
        Console.Write($"{col}\t");
    }
    
    Console.Write("\n");
}
````

## Batch Updates
Google Sheets API has a method to perform batch updates.  The purpose is to enable you to update mulitple values in a single operation and also avoid hitting the API throttling limits.  ````GoogleSheetsWrapper```` lets you use this API.  One important thing to understand is you need to specify which properties you want updated with the field mask paramater.  A few examples are below,

### Update UserEnteredValue and UserEnteredFormat

````csharp
var updates = new List<BatchUpdateRequestObject>();
updates.Add(new BatchUpdateRequestObject()
{
    Range = new SheetRange("A8:B8"),
    Data = new CellData()
    {
        UserEnteredValue = new ExtendedValue()
        {
            StringValue = "Hello World"
        },
        UserEnteredFormat = new CellFormat()
        {
            BackgroundColor = new Color()
            {
                Blue = 0,
                Green = 1,
                Red = 0,
            }
        }
    }
});

// Note the field mask is for both the UserEnteredValue and the UserEnteredFormat below,
sheetHelper.BatchUpdate(updates, "userEnteredValue, userEnteredFormat");
````

### Update only UserEnteredValue
The ````GoogleSheetsWrapper```` library defaults to only updating the ````UserEnteredValue````.  This is to allow you to keep your existing cell formating in place.  However, we give you the option like above to override that behavior.

````csharp
var updates = new List<BatchUpdateRequestObject>();
updates.Add(new BatchUpdateRequestObject()
{
    Range = new SheetRange("A8:B8"),
    Data = new CellData()
    {
        UserEnteredValue = new ExtendedValue()
        {
            StringValue = "Hello World"
        }
    }
});

// Note the field mask parameter not being specified here defaults to => "userEnteredValue"
sheetHelper.BatchUpdate(updates);
````
### Update all Cell Properties
The Google Sheets API lets you use the ````*```` wildcard to specify all properties to be updated.  Note when you do this, if any values are null / empty, this will clear out those settings when you save them!  Example below,

````csharp
var updates = new List<BatchUpdateRequestObject>();
updates.Add(new BatchUpdateRequestObject()
{
    Range = new SheetRange("A8:B8"),
    Data = new CellData()
    {
        UserEnteredValue = new ExtendedValue()
        {
            StringValue = "Hello World"
        }
    }
});

// Note setting the field mask to "*" tells the API to save all property values, even if they are null / blank
sheetHelper.BatchUpdate(updates, "*");
````

### All Batch Operation Field Mask Options

Full list of the different fields you can specify in the field mask property are below.  To combine any of these you want to use a comma separated string.

| C# Property Name | Google Sheets API Property Name |
| ---------------- | ------------------------------- |
| ````DataSourceFormula```` | ````dataSourceFormula```` |
| ````DataSourceTable```` | ````dataSourceTable```` |
| ````DataValidation```` | ````dataValidation```` |
| ````EffectiveFormat```` | ````effectiveFormat```` |
| ````EffectiveValue```` | ````effectiveValue```` |
| ````FormattedValue```` | ````formattedValue```` |
| ````Hyperlink```` | ````hyperlink```` |
| ````Note```` | ````note```` |
| ````PivotTable```` | ````pivotTable```` |
| ````TextFormatRuns```` | ````textFormatRuns```` |
| ````UserEnteredFormat```` | ````userEnteredFormat```` |
| ````UserEnteredValue```` | ````userEnteredValue```` |

## Append a CSV to Google Sheets

```csharp
var appender = new SheetAppender(
    // https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID_IS_HERE>/edit#gid=0
    settings.GoogleSpreadsheetId,
    // The email for the service account you created
    settings.GoogleServiceAccountName,
    // the name of the tab you want to access, leave blank if you want the default first tab
    settings.GoogleMainSheetName);

appender.Init(settings.JsonCredential);

var filepath = @"C:\Input\input.csv";

using (var stream = new FileStream(filepath, FileMode.Open))
{
    // Append the csv file to Google sheets, include the header row 
    // and wait 1000 milliseconds between batch updates 
    // Google Sheets API throttles requests per minute so you may need to play
    // with this setting. 
    appender.AppendCsv(
        stream, // The CSV FileStrem 
        true, // true indicating to include the header row
        1000); // 1000 milliseconds to wait every 100 rows that are batch sent to the Google Sheets API
}

```
## Import a CSV to Google Sheets and Purge Existing Rows

If you wanted to delete the existing rows in your tab first, you can use
```SheetHelper``` methods to do that.  Below is a sample of that,

```csharp
// Create a SheetHelper class
var sheetHelper = new SheetHelper(
    // https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID_IS_HERE>/edit#gid=0
    settings.GoogleSpreadsheetId,
    // The email for the service account you created
    settings.GoogleServiceAccountName,
    // the name of the tab you want to access, leave blank if you want the default first tab
    settings.GoogleMainSheetName);

sheetHelper.Init(settings.JsonCredential);

// Get the total row count for the existing sheet
var rows = sheetHelper.GetRows(new SheetRange("", 1, 1, 1));

// Delete all of the rows
sheetHelper.DeleteRows(1, rows.Count);

// Create the SheetAppender class
var appender = new SheetAppender(sheetHelper);

var filepath = @"C:\Input\input.csv";

using (var stream = new FileStream(filepath, FileMode.Open))
{
    // Append the csv file to Google sheets, include the header row 
    // and wait 1000 milliseconds between batch updates 
    // Google Sheets API throttles requests per minute so you may need to play
    // with this setting. 
    appender.AppendCsv(
        stream, // The CSV FileStrem 
        true, // true indicating to include the header row
        1000); // 1000 milliseconds to wait every 100 rows that are batch sent to the Google Sheets API
}

```

## Export Google Sheet to CSV

```csharp
// Create a SheetHelper class
var sheetHelper = new SheetHelper(
    // https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID_IS_HERE>/edit#gid=0
    settings.GoogleSpreadsheetId,
    // The email for the service account you created
    settings.GoogleServiceAccountName,
    // the name of the tab you want to access, leave blank if you want the default first tab
    settings.GoogleMainSheetName);

sheetHelper.Init(settings.JsonCredential);

var exporter = new SheetExporter(sheetHelper);

var filepath = @"C:\Output\output.csv";


// OPTION 1: Default to CultureInfo.InvariantCulture and "," as the delimiter
using (var stream = new FileStream(filepath, FileMode.Create))
{
    // Query the range A1:G (ie 1st column, 1st row, 8th column and last row in the sheet)
    var range = new SheetRange("TAB_NAME", 1, 1, 8);
    exporter.ExportAsCsv(range, stream);
}

// OPTION 2: Create your own CsvConfiguration object with full control on Culture and delimiter
using (var stream = new FileStream(filepath, FileMode.Create))
{
    // Query the range A1:G (ie 1st column, 1st row, 8th column and last row in the sheet)
    var range = new SheetRange("TAB_NAME", 1, 1, 8);

    // This gives you full control on the CSV file CultureInfo and Delimiter value
    var csvConfig = new CsvConfiguration(CultureInfo.CurrentCulture)
    {
        Delimiter = ";",
    };

    // Export with the customer CsvConfiguration value
    exporter.ExportAsCsv(range, stream, csvConfig);
}

```

## Export Google Sheet to Excel File

```csharp
var exporter = new SheetExporter(
    // https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID_IS_HERE>/edit#gid=0
    settings.GoogleSpreadsheetId,
    // The email for the service account you created
    settings.GoogleServiceAccountName,
    // the name of the tab you want to access, leave blank if you want the default first tab
    settings.GoogleMainSheetName);

exporter.Init(settings.JsonCredential);

var filepath = @"C:\Output\output.xlsx";

using (var stream = new FileStream(filepath, FileMode.Create))
{
    // Query the range A1:G (ie 1st column, 1st row, 8th column and last row in the sheet)
    var range = new SheetRange("TAB_NAME", 1, 1, 8);
    exporter.ExportAsExcel(range, stream);
}

```

## Authentication
> [!IMPORTANT] 
> You need to setup a Google API Service Account before you can use this library.  

1. If you have not yet created a Google Cloud project, you will need to create one before you create a service account.  Documentation on this can be found below,

    https://developers.google.com/workspace/guides/create-project

2. Create a service account.  Steps to do that are documented below,

    https://cloud.google.com/docs/authentication/production#create_service_account

3. Enable the ````Google Sheets API```` for your Google Cloud project.  Details on this can be found below,

    https://cloud.google.com/endpoints/docs/openapi/enable-api
  
4. For the service account you created, create a new service account key.  When you do this, it will also download the key, choose the JSON key format.

    https://cloud.google.com/iam/docs/keys-create-delete  

5. Use the service account identity that is created and add that email address to grant it permissions to the Google Sheets Spreadsheet you want to interact with.

6. Configure your code with the following parameters to initialize a ````SheetHelper```` object

```csharp
// You need to implement your own configuration management solution here!
var settings = AppSettings.FromEnvironment();

// Create a SheetHelper class for the specified Google Sheet and Tab name
var sheetHelper = new SheetHelper<TestRecord>(
    // https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID_IS_HERE>/edit#gid=0
    settings.GoogleSpreadsheetId,
    // The email for the service account you created
    settings.GoogleServiceAccountName,
    // the name of the tab you want to access, leave blank if you want the default first tab
    settings.GoogleMainSheetName);

sheetHelper.Init(settings.JsonCredential);
```

Another good article on how to setup a Google Service Account can be found below on Robocorp's documentation site,

https://robocorp.com/docs/development-guide/google-sheets/interacting-with-google-sheets#create-a-google-service-account
