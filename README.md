# GoogleSheetsWrapper
## Google Sheets API .NET Wrapper Library

This library allows you to use strongly typed objects against a Google Sheets spreadsheet without having to have knowledge on the Google Sheets API methods and protocols. 

The following Google Sheets API operations are supported: 

* Reading all rows
* Appending new rows
* Deleting rows
* Updating specific cells

All operations above are encapsulated in the SheetHelper class. 

There are also base classes, BaseRecord and BaseRepository to simply transforming raw Google Sheets rows into .NET objects. 

Extending the BaseRecord class you can decorate properties with the SheetFieldAttribute to describe the column header name, the column index (1 based index and not 0 based index)

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
    public long PhoneNumber { get; set; }

    [SheetField(
        DisplayName = "Previous Donation Amount",
        ColumnID = 3,
        FieldType = SheetFieldType.Currency)]
    public double DonationAmount { get; set; }

    [SheetField(
        DisplayName = "Date",
        ColumnID = 4,
        FieldType = SheetFieldType.DateTime)]
    public DateTime DateTime { get; set; }

    public TestRecord() { }

    public TestRecord(IList<object> row, int rowId) : base(row, rowId) { }
}
```

Extending the BaseRepository allows you to define your own access layer to the Google Sheets tab you want to work with. 

```csharp
public class TestRepository : BaseRepository<TestRecord>
{
    public TestRepository() { }

    public TestRepository(SheetHelper<TestRecord> sheetsHelper)
        : base(sheetsHelper) { }
}
```

Using the Google Sheets Wrapper library.  

>***Note** that to do this you will need to setup a Google Service Account and create a service account key.  More details can be found here: 
https://cloud.google.com/iam/docs/creating-managing-service-account-keys#creating_service_account_keys*

```csharp
// You need to implement your own configuration management solution here!
var settings = AppSettings.FromEnvironment();

// Create a SheetHelper class for the specified Google Sheet and Tab name
var sheetHelper = new SheetHelper<TestRecord>(
    settings.GoogleSpreadsheetId,
    settings.GoogleServiceAccountName,
    settings.GoogleMainSheetName);

sheetHelper.Init(settings.JsonCredential);

// Create a Repository for the TestRecord class
var repository = new TestRepository(sheetHelper);

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

// Update the DonationAmount field and save it back to Google Sheets
firstRecord.DonationAmount = 99.99;
repository.SaveField(firstRecord, (r) => r.DonationAmount);

// Delete the first record from Google Sheets
repository.DeleteRecord(firstRecord);

// Add a new record to the end of the Google Sheets tab
repository.AddRecord(new TestRecord()
{
    Name = "John",
    Number = 2021112222,
    DonationAmount = 250.50,
    Date = DateTime.Now()
});

```
