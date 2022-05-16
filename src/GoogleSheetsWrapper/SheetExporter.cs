﻿using CsvHelper;
using CsvHelper.Configuration;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace GoogleSheetsWrapper
{
    public class SheetExporter
    {
        private SheetHelper _sheetHelper;

        public SheetExporter(string spreadsheetID, string serviceAccountEmail, string tabName)
        {
            this._sheetHelper = new SheetHelper(spreadsheetID, serviceAccountEmail, tabName);
        }

        public SheetExporter(SheetHelper sheetHelper)
        {
            this._sheetHelper = sheetHelper;
        }

        public void Init(string jsonCredentials)
        {
            this._sheetHelper.Init(jsonCredentials);
        }

        /// <summary>
        /// Exports the current Google Sheet tab to a CSV file
        /// </summary>
        /// <param name="range"></param>
        /// <param name="stream"></param>
        public void ExportAsCsv(SheetRange range, Stream stream, string delimiter = ",")
        {
            var rows = this._sheetHelper.GetRowsFormatted(range);

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = delimiter,
            };

            using StreamWriter streamWriter = new StreamWriter(stream);
            using var csv = new CsvWriter(streamWriter, config);
            foreach (var row in rows)
            {
                foreach (var cell in row)
                {
                    csv.WriteField(cell?.ToString());
                }

                csv.NextRecord();
            }
        }

        /// <summary>
        /// Exports the current Google Sheet tab to a CSV file
        /// </summary>
        /// <param name="range"></param>
        /// <param name="stream"></param>
        public void ExportAsExcel(SheetRange range, Stream stream)
        {
            var rows = this._sheetHelper.GetRowsFormatted(range);

            SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = document.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            // Add Sheets to the Workbook.
            Sheets sheets = document.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = string.IsNullOrEmpty(this._sheetHelper.TabName) ? "Sheet1" : this._sheetHelper.TabName
            };
            sheets.Append(sheet);

            foreach (var row in rows)
            {
                var newRow = new Row();

                foreach (var value in row)
                {
                    Cell cell = this.CreateCell(value?.ToString());
                    newRow.AppendChild(cell);
                }

                sheetData.AppendChild(newRow);
            }

            workbookpart.Workbook.Save();

            document.Close();
        }

        private Cell CreateCell(string text)
        {
            Cell cell = new Cell
            {
                DataType = ResolveCellDataTypeOnValue(text),
                CellValue = new CellValue(text)
            };

            return cell;
        }

        private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }
    }
}
