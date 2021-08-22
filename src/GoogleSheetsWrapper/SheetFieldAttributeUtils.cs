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
        public static void PopulateRecord<T>(T record, IList<object> row, int minColumnId = 1) where T : BaseRecord
        {
            var properties = record.GetType().GetProperties();

            foreach (var property in properties)
            {
                var attribute = property.GetCustomAttribute<SheetFieldAttribute>();

                if ((attribute != null) && (row.Count > (attribute.ColumnID - minColumnId)))
                {
                    var stringValue = row[attribute.ColumnID - minColumnId]?.ToString();

                    if (attribute.FieldType == SheetFieldType.String)
                    {
                        property.SetValue(record, stringValue);
                    }
                    else if (attribute.FieldType == SheetFieldType.Currency)
                    {
                        if (!string.IsNullOrWhiteSpace(stringValue))
                        {
                            var value = CurrencyParsing.ParseCurrencyString(stringValue);
                            property.SetValue(record, value);
                        }
                    }
                    else if (attribute.FieldType == SheetFieldType.PhoneNumber)
                    {
                        var value = PhoneNumberParsing.RemoveUSInterationalPhoneCode(stringValue);

                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            property.SetValue(record, long.Parse(value));
                        }
                    }
                    else if (attribute.FieldType == SheetFieldType.DateTime)
                    {
                        DateTime? dt = new DateTime();

                        if (!string.IsNullOrWhiteSpace(stringValue))
                        {
                            var serialNumber = double.Parse(stringValue);

                            dt = DateTimeUtils.ConvertFromSerialNumber(serialNumber);

                            property.SetValue(record, dt);
                        }
                    }
                    else if (attribute.FieldType == SheetFieldType.Number)
                    {
                        if (!string.IsNullOrWhiteSpace(stringValue))
                        {
                            var value = double.Parse(stringValue);

                            property.SetValue(record, value);
                        }
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

            var value = property.GetValue(record);

            if (attribute.FieldType == SheetFieldType.String)
            {
                if (value == null)
                {
                    cell.UserEnteredValue.StringValue = string.Empty;
                }
                else if (string.IsNullOrWhiteSpace(value.ToString()))
                {
                    cell.UserEnteredValue.StringValue = string.Empty;
                }
                else
                {
                    cell.UserEnteredValue.StringValue = value.ToString();
                }
            }
            else if (attribute.FieldType == SheetFieldType.Currency)
            {
                if (value != null)
                {
                    cell.UserEnteredValue.NumberValue = double.Parse(value.ToString());
                }

                cell.UserEnteredFormat = new CellFormat()
                {
                    NumberFormat = new NumberFormat()
                    {
                        Pattern = attribute.NumberFormatPattern,
                        Type = "CURRENCY"
                    }
                };

            }
            else if (attribute.FieldType == SheetFieldType.PhoneNumber)
            {
                if (value != null)
                {
                    double parsedNumber = double.Parse(value.ToString());

                    if (parsedNumber != 0)
                    {
                        cell.UserEnteredValue.NumberValue = parsedNumber;
                    }
                    else
                    {
                        cell.UserEnteredValue.NumberValue = null;
                    }
                }

                cell.UserEnteredFormat = new CellFormat()
                {
                    NumberFormat = new NumberFormat()
                    {
                        Pattern = attribute.NumberFormatPattern,
                        Type = "NUMBER"
                    }
                };
            }
            else if (attribute.FieldType == SheetFieldType.DateTime)
            {
                if (value != null)
                {
                    cell.UserEnteredValue.NumberValue = DateTimeUtils.ConvertToSerialNumber((DateTime)value);
                }

                cell.UserEnteredFormat = new CellFormat()
                {
                    NumberFormat = new NumberFormat()
                    {
                        Pattern = attribute.NumberFormatPattern,
                        Type = "NUMBER"
                    }
                };
            }
            else if (attribute.FieldType == SheetFieldType.Number)
            {
                if (value != null)
                {
                    cell.UserEnteredValue.NumberValue = double.Parse(value.ToString());
                }

                cell.UserEnteredFormat = new CellFormat()
                {
                    NumberFormat = new NumberFormat()
                    {
                        Pattern = attribute.NumberFormatPattern,
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
