using GoogleSheetsWrapper.Utils;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace GoogleSheetsWrapper
{
    public class SheetFieldAttributeUtils
    {
        public static void PopulateRecord<T>(T record, IList<object> row) where T : BaseRecord
        {
            var properties = record.GetType().GetProperties();

            foreach (var property in properties)
            {
                var attribute = property.GetCustomAttribute<SheetFieldAttribute>();

                if ((attribute != null) && (row.Count >= attribute.ColumnID))
                {
                    var stringValue = row[attribute.ColumnID - 1]?.ToString();

                    if (attribute.FieldType == SheetFieldType.String)
                    {
                        property.SetValue(record, stringValue);
                    }
                    else if (attribute.FieldType == SheetFieldType.Currency)
                    {
                        var value = CurrencyParsing.ParseCurrencyString(stringValue);
                        property.SetValue(record, value);
                    }
                    else if (attribute.FieldType == SheetFieldType.PhoneNumber)
                    {
                        var value = PhoneNumberParsing.RemoveUSInterationalPhoneCode(stringValue);
                        property.SetValue(record, long.Parse(value));
                    }
                    else if (attribute.FieldType == SheetFieldType.DateTime)
                    {
                        var serialNumber = double.Parse(stringValue);

                        DateTime dt = DateTimeUtils.ConvertFromSerialNumber(serialNumber);

                        property.SetValue(record, dt);
                    }
                    else
                    {
                        throw new ArgumentException($"{attribute.FieldType} not supported yet!");
                    }
                }
            }
        }

        public static CellData GetCellDataForSheetField<T>(T record, SheetFieldAttribute attribute, PropertyInfo property)
        {
            var cell = new CellData
            {
                UserEnteredValue = new ExtendedValue()
            };

            if (attribute.FieldType == SheetFieldType.String)
            {
                cell.UserEnteredValue.StringValue = property.GetValue(record)?.ToString();
            }
            else if (attribute.FieldType == SheetFieldType.Currency)
            {
                cell.UserEnteredValue.NumberValue = double.Parse(property.GetValue(record).ToString());
                cell.UserEnteredFormat = new CellFormat()
                {
                    NumberFormat = new NumberFormat()
                    {
                        Pattern = "\"$\"#,##0.00",
                        Type = "CURRENCY"
                    }
                };
            }
            else if (attribute.FieldType == SheetFieldType.PhoneNumber)
            {
                cell.UserEnteredValue.NumberValue = double.Parse(property.GetValue(record).ToString());
                cell.UserEnteredFormat = new CellFormat()
                {
                    NumberFormat = new NumberFormat()
                    {
                        Pattern = "(###)\" \"###\"-\"####",
                        Type = "NUMBER"
                    }
                };
            }
            else if (attribute.FieldType == SheetFieldType.DateTime)
            {
                cell.UserEnteredValue.NumberValue = DateTimeUtils.ConvertToSerialNumber((DateTime)property.GetValue(record));
                cell.UserEnteredFormat = new CellFormat()
                {
                    NumberFormat = new NumberFormat()
                    {
                        Pattern = "M/d/yyyy H:mm:ss",
                        Type = "NUMBER"
                    }
                };
            }
            else
            {
                throw new ArgumentException($"{attribute.FieldType} is not supported yet!");
            }

            return cell;
        }

        public static int GetColumnId<T>(Expression<Func<T, object>> expression) where T : BaseRecord
        {
            var attribute = GetSheetFieldAttribute(expression);

            return attribute.ColumnID;
        }

        public static SortedDictionary<SheetFieldAttribute, PropertyInfo> GetAllSheetFieldAttributes<T>()
        {
            return GetAllSheetFieldAttributes(typeof(T));
        }

        public static SortedDictionary<SheetFieldAttribute, PropertyInfo> GetAllSheetFieldAttributes(Type type)
        {
            var result = new SortedDictionary<SheetFieldAttribute, PropertyInfo>(new SheetFieldAttributeComparer());

            var properties = type.GetProperties();

            foreach (var property in properties)
            {
                var attribute = property.GetCustomAttribute<SheetFieldAttribute>();

                if (attribute != null)
                {
                    result.Add(attribute, property);
                }
            }

            return result;
        }

        public static SheetFieldAttribute GetSheetFieldAttribute<T>
            (Expression<Func<T, object>> expression)
        {
            MemberExpression memberExpression;

            if (expression.Body is MemberExpression)
            {
                memberExpression = (MemberExpression)expression.Body;
            }
            else if (expression.Body is UnaryExpression unaryExpression)
            {
                memberExpression = (MemberExpression)unaryExpression.Operand;
            }
            else
            {
                throw new ArgumentException();
            }

            var propertyInfo = (PropertyInfo)memberExpression.Member;
            return propertyInfo.GetCustomAttribute<SheetFieldAttribute>();
        }
    }
}
